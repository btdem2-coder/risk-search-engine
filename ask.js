require("dotenv").config();

const fs = require("fs");
const axios = require("axios");
const OpenAI = require("openai");
const readlineSync = require("readline-sync");

const {
  AZURE_OPENAI_ENDPOINT,
  AZURE_OPENAI_KEY,
  AZURE_OPENAI_DEPLOYMENT,
  AZURE_SEARCH_ENDPOINT,
  AZURE_SEARCH_KEY,
  AZURE_SEARCH_INDEX,
  OPENAI_API_KEY,
  OPENAI_FILTER_MODEL,
} = process.env;

const openai = new OpenAI({ apiKey: OPENAI_API_KEY });
const records = JSON.parse(fs.readFileSync("./records.json", "utf8"));

function uniqueValues(field) {
  return [...new Set(records.map((r) => r[field]).filter(Boolean))].slice(0, 300);
}

function escapeOData(value) {
  return String(value).replace(/'/g, "''");
}

async function extractFilters(question) {
  const industries = uniqueValues("industry");
  const entities = uniqueValues("entity");

  console.log("Available industries:", industries);
  console.log("Available entities:", entities);

  const response = await openai.chat.completions.create({
    model: OPENAI_FILTER_MODEL || "gpt-4o-mini",
    temperature: 0,
    messages: [
      {
        role: "system",
        content:
          "Extract exact filters from the user's question. Return arrays for industry and entity when multiple exact matches exist, otherwise return a single-item array. Only use values from the provided allowed lists. If no exact match exists for a field, return null. Return JSON only.",
      },
      {
        role: "user",
        content: JSON.stringify({
          question,
          allowedIndustries: industries,
          allowedEntities: entities,
          outputFormat: {
            industry: "string[]|null",
            entity: "string[]|null",
          },
        }),
      },
    ],
  });

  const raw = response.choices[0].message.content
  .replace(/```json/g, "")
  .replace(/```/g, "")
  .trim();
  
  return normalizeFilters(JSON.parse(raw));
}

function normalizeFilters(filters = {}) {
  const toArrayOrNull = (value) => {
    if (Array.isArray(value)) {
      const cleaned = value
        .map((item) => String(item || "").trim())
        .filter(Boolean);
      return cleaned.length ? cleaned : null;
    }
    if (value === null || value === undefined) return null;
    const cleaned = String(value).trim();
    return cleaned ? [cleaned] : null;
  };

  return {
    industry: toArrayOrNull(filters.industry),
    entity: toArrayOrNull(filters.entity),
  };
}

function findMatchingValues(values, allowedValues, { exact = false } = {}) {
  if (!values || !values.length) return [];

  const matches = new Set();

  values.forEach((value) => {
    const v = String(value).toLowerCase();
    allowedValues.forEach((item) => {
      const i = item.toLowerCase();
      const isMatch = exact ? i === v : i.includes(v) || v.includes(i);
      if (isMatch) {
        matches.add(item);
      }
    });
  });

  return [...matches];
}

function expandEntityFamiliesFromQuestion(question, filters) {
  const normalizedFilters = normalizeFilters(filters);
  const extractedEntities = normalizedFilters.entity || [];
  if (!extractedEntities.length) return normalizedFilters;

  const questionLower = String(question || "").toLowerCase();
  const allowedEntities = uniqueValues("entity");
  const expanded = new Set();

  extractedEntities.forEach((entity) => {
    const basePhrase = String(entity).split(" - ")[0].trim().toLowerCase();
    if (!basePhrase) return;

    // Only expand for phrases actually present in the prompt.
    if (!questionLower.includes(basePhrase)) return;

    allowedEntities.forEach((allowedEntity) => {
      const candidate = allowedEntity.toLowerCase();
      if (candidate.startsWith(basePhrase) || candidate.includes(basePhrase)) {
        expanded.add(allowedEntity);
      }
    });
  });

  if (!expanded.size) {
    return normalizedFilters;
  }

  return {
    ...normalizedFilters,
    entity: [...expanded],
  };
}

function expandIndustriesFromEntities(filters) {
  const normalizedFilters = normalizeFilters(filters);
  const entities = normalizedFilters.entity || [];
  if (!entities.length) return normalizedFilters;

  const entitySet = new Set(entities);
  const linkedIndustries = new Set(normalizedFilters.industry || []);

  records.forEach((record) => {
    if (!record.entity || !record.industry) return;
    if (entitySet.has(record.entity)) {
      linkedIndustries.add(record.industry);
    }
  });

  return {
    ...normalizedFilters,
    industry: linkedIndustries.size ? [...linkedIndustries] : normalizedFilters.industry,
  };
}

async function generateEmbedding(text) {
  const url = `${AZURE_OPENAI_ENDPOINT}/openai/deployments/${AZURE_OPENAI_DEPLOYMENT}/embeddings?api-version=2023-05-15`;

  const response = await axios.post(
    url,
    { input: text },
    {
      headers: {
        "Content-Type": "application/json",
        "api-key": AZURE_OPENAI_KEY,
      },
    }
  );

  return response.data.data[0].embedding;
}

async function searchAzure(question, filters) {
  const trimmedQuestion = String(question || "").trim();
  const normalizedFilters = normalizeFilters(filters);
  const useExactFilterMatching = !trimmedQuestion;

  const filterParts = [];
  const industries = uniqueValues("industry");
  const entities = uniqueValues("entity");
  const matchedIndustries = findMatchingValues(normalizedFilters.industry, industries, {
    exact: useExactFilterMatching,
  });
  const matchedEntities = findMatchingValues(normalizedFilters.entity, entities, {
    exact: useExactFilterMatching,
  });

  if (matchedIndustries.length) {
    filterParts.push(
      "(" +
        matchedIndustries
          .map((v) => `industry eq '${escapeOData(v)}'`)
          .join(" or ") +
        ")"
    );
  }
  
  if (matchedEntities.length) {
    filterParts.push(
      "(" +
        matchedEntities
          .map((v) => `entity eq '${escapeOData(v)}'`)
          .join(" or ") +
        ")"
    );
  }

  const body = { count: true };

  if (trimmedQuestion) {
    const embedding = await generateEmbedding(trimmedQuestion);
    body.vectorQueries = [
      {
        kind: "vector",
        vector: embedding,
        fields: "embedding",
        k: 10,
      },
    ];
  } else {
    body.search = "*";
    body.top = 1000;
  }

  if (filterParts.length) {
    body.filter = filterParts.join(" and ");
  }

  console.log("REQUEST BODY:", JSON.stringify(body, null, 2));


  const url = `${AZURE_SEARCH_ENDPOINT}/indexes/${AZURE_SEARCH_INDEX}/docs/search?api-version=2023-11-01`;

  const response = await axios.post(url, body, {
    headers: {
      "Content-Type": "application/json",
      "api-key": AZURE_SEARCH_KEY,
    },
  });

  console.log("RAW RESPONSE:", JSON.stringify(response.data, null, 2));

  return response.data.value;
}

function hydrate(searchResults) {
  return searchResults.map((result) => {
    const full = records.find((r) => r.id === result.id);
    return {
      score: result["@search.score"],
      id: result.id,
      ...full,
    };
  });
}

function getFilterOptions() {
  const linksSet = new Set();
  records.forEach((record) => {
    if (!record.industry || !record.entity) return;
    linksSet.add(
      JSON.stringify({
        industry: record.industry,
        entity: record.entity,
      })
    );
  });

  return {
    industries: uniqueValues("industry"),
    entities: uniqueValues("entity"),
    links: [...linksSet].map((item) => JSON.parse(item)),
  };
}

async function askRiskSearch(question, selectedFilters = {}) {
  const trimmedQuestion = String(question || "").trim();
  const normalizedSelectedFilters = normalizeFilters(selectedFilters);
  const hasManualFilters =
    Boolean(normalizedSelectedFilters.industry?.length) ||
    Boolean(normalizedSelectedFilters.entity?.length);

  if (!trimmedQuestion && !hasManualFilters) {
    throw new Error("Question is required.");
  }

  const filters = hasManualFilters
    ? normalizedSelectedFilters
    : expandIndustriesFromEntities(
        expandEntityFamiliesFromQuestion(
          trimmedQuestion,
          await extractFilters(trimmedQuestion)
        )
      );

  const results = await searchAzure(trimmedQuestion, filters);
  const hydrated = hydrate(results);

  return { filters, results: hydrated };
}

async function main() {
  console.log("\nRisk Search Console\n");

  const question = readlineSync.question("Ask your question: ");

  console.log("\nExtracting filters...");
  const filters = await extractFilters(question);
  console.log("Filters:", filters);

  console.log("\nSearching Azure AI Search...");
  const results = await searchAzure(question, filters);
  const hydrated = hydrate(results);

  console.log("\nTop risks:\n");

  hydrated.forEach((risk, index) => {
    console.log(`${index + 1}. ${risk.riskSubject || risk.id}`);
    console.log(`Score: ${risk.score}`);
    console.log(`Industry: ${risk.industry || ""}`);
    console.log(`Entity: ${risk.entity || ""}`);
    console.log(`Description: ${risk.riskDescription || ""}`);
    console.log("-".repeat(80));
  });
}

module.exports = {
  askRiskSearch,
  getFilterOptions,
  extractFilters,
  searchAzure,
  hydrate,
};

if (require.main === module) {
  main().catch((error) => {
    console.error("Error:", error.response?.data || error.message);
  });
}