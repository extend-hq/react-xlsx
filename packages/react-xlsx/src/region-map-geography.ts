import { feature as topojsonFeature } from "topojson-client";
import countiesAlbers10m from "us-atlas/counties-albers-10m.json";
import countries50m from "world-atlas/countries-50m.json";
import type { Feature, FeatureCollection, Geometry } from "geojson";

export { geoIdentity, geoMercator, geoNaturalEarth1, geoPath } from "d3-geo";

export type RegionMapFeature = Feature<Geometry, { name?: string; regionSet?: "country" | "us-state"; stateCode?: string }>;

const WORLD_COUNTRY_FEATURES = ((topojsonFeature(
  countries50m as unknown as Parameters<typeof topojsonFeature>[0],
  (countries50m as { objects: { countries: unknown } }).objects.countries as Parameters<typeof topojsonFeature>[1]
) as unknown) as FeatureCollection<Geometry, { name?: string }>).features as RegionMapFeature[];

const US_STATE_NAME_BY_ID: Record<string, { code: string; name: string }> = {
  "01": { code: "AL", name: "Alabama" },
  "02": { code: "AK", name: "Alaska" },
  "04": { code: "AZ", name: "Arizona" },
  "05": { code: "AR", name: "Arkansas" },
  "06": { code: "CA", name: "California" },
  "08": { code: "CO", name: "Colorado" },
  "09": { code: "CT", name: "Connecticut" },
  "10": { code: "DE", name: "Delaware" },
  "11": { code: "DC", name: "District of Columbia" },
  "12": { code: "FL", name: "Florida" },
  "13": { code: "GA", name: "Georgia" },
  "15": { code: "HI", name: "Hawaii" },
  "16": { code: "ID", name: "Idaho" },
  "17": { code: "IL", name: "Illinois" },
  "18": { code: "IN", name: "Indiana" },
  "19": { code: "IA", name: "Iowa" },
  "20": { code: "KS", name: "Kansas" },
  "21": { code: "KY", name: "Kentucky" },
  "22": { code: "LA", name: "Louisiana" },
  "23": { code: "ME", name: "Maine" },
  "24": { code: "MD", name: "Maryland" },
  "25": { code: "MA", name: "Massachusetts" },
  "26": { code: "MI", name: "Michigan" },
  "27": { code: "MN", name: "Minnesota" },
  "28": { code: "MS", name: "Mississippi" },
  "29": { code: "MO", name: "Missouri" },
  "30": { code: "MT", name: "Montana" },
  "31": { code: "NE", name: "Nebraska" },
  "32": { code: "NV", name: "Nevada" },
  "33": { code: "NH", name: "New Hampshire" },
  "34": { code: "NJ", name: "New Jersey" },
  "35": { code: "NM", name: "New Mexico" },
  "36": { code: "NY", name: "New York" },
  "37": { code: "NC", name: "North Carolina" },
  "38": { code: "ND", name: "North Dakota" },
  "39": { code: "OH", name: "Ohio" },
  "40": { code: "OK", name: "Oklahoma" },
  "41": { code: "OR", name: "Oregon" },
  "42": { code: "PA", name: "Pennsylvania" },
  "44": { code: "RI", name: "Rhode Island" },
  "45": { code: "SC", name: "South Carolina" },
  "46": { code: "SD", name: "South Dakota" },
  "47": { code: "TN", name: "Tennessee" },
  "48": { code: "TX", name: "Texas" },
  "49": { code: "UT", name: "Utah" },
  "50": { code: "VT", name: "Vermont" },
  "51": { code: "VA", name: "Virginia" },
  "53": { code: "WA", name: "Washington" },
  "54": { code: "WV", name: "West Virginia" },
  "55": { code: "WI", name: "Wisconsin" },
  "56": { code: "WY", name: "Wyoming" }
};

const US_STATE_FEATURES = ((topojsonFeature(
  countiesAlbers10m as unknown as Parameters<typeof topojsonFeature>[0],
  (countiesAlbers10m as { objects: { states: unknown } }).objects.states as Parameters<typeof topojsonFeature>[1]
) as unknown) as FeatureCollection<Geometry, { name?: string }>).features.map((feature) => {
  const id = typeof feature.id === "string" ? feature.id : String(feature.id ?? "");
  const state = US_STATE_NAME_BY_ID[id];
  return {
    ...feature,
    properties: {
      ...(feature.properties ?? {}),
      name: state?.name,
      regionSet: "us-state" as const,
      stateCode: state?.code
    }
  };
}) as RegionMapFeature[];

const REGION_MAP_COUNTRY_ALIASES = new Map<string, string>([
  ["us", "united states of america"],
  ["usa", "united states of america"],
  ["u s a", "united states of america"],
  ["united states", "united states of america"],
  ["united states america", "united states of america"],
  ["u s", "united states of america"],
  ["uk", "united kingdom"],
  ["u k", "united kingdom"],
  ["uae", "united arab emirates"],
  ["u a e", "united arab emirates"],
  ["south korea", "korea, south"],
  ["north korea", "korea, north"],
  ["russia", "russian federation"],
  ["vietnam", "viet nam"],
  ["czech republic", "czechia"],
  ["ivory coast", "cote d'ivoire"],
  ["côte divoire", "cote d'ivoire"]
]);

const REGION_MAP_US_STATE_ALIASES = new Map<string, string>([
  ["district of columbia", "district of columbia"],
  ["washington dc", "district of columbia"],
  ["washington d c", "district of columbia"],
  ["dc", "district of columbia"],
  ["d c", "district of columbia"]
]);

const REGION_MAP_FEATURES_BY_KEY = (() => {
  const byKey = new Map<string, RegionMapFeature>();
  WORLD_COUNTRY_FEATURES.forEach((feature) => {
    const name = typeof feature.properties?.name === "string" ? feature.properties.name : "";
    const key = normalizeRegionMapKey(name);
    if (key.length > 0) {
      byKey.set(key, feature);
    }
  });
  return byKey;
})();

const REGION_MAP_US_STATE_FEATURES_BY_KEY = (() => {
  const byKey = new Map<string, RegionMapFeature>();
  US_STATE_FEATURES.forEach((feature) => {
    const name = typeof feature.properties?.name === "string" ? feature.properties.name : "";
    const key = normalizeRegionMapKey(name);
    if (key.length > 0) {
      byKey.set(key, feature);
    }
    const stateCode = normalizeRegionMapKey(feature.properties?.stateCode);
    if (stateCode.length > 0) {
      byKey.set(stateCode, feature);
    }
  });
  REGION_MAP_US_STATE_ALIASES.forEach((value, key) => {
    const feature = byKey.get(normalizeRegionMapKey(value));
    if (feature) {
      byKey.set(normalizeRegionMapKey(key), feature);
    }
  });
  return byKey;
})();

function normalizeRegionMapKey(value: unknown) {
  if (typeof value !== "string") {
    return "";
  }
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/&/g, " and ")
    .replace(/[^a-zA-Z0-9]+/g, " ")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}


export function resolveRegionMapFeature(value: unknown, featureSet: "country" | "us-state" = "country") {
  const rawKey = normalizeRegionMapKey(value);
  if (!rawKey) {
    return null;
  }
  if (featureSet === "us-state") {
    const canonicalKey = REGION_MAP_US_STATE_ALIASES.get(rawKey) ?? rawKey;
    return REGION_MAP_US_STATE_FEATURES_BY_KEY.get(canonicalKey) ?? null;
  }
  const canonicalKey = REGION_MAP_COUNTRY_ALIASES.get(rawKey) ?? rawKey;
  return REGION_MAP_FEATURES_BY_KEY.get(canonicalKey) ?? null;
}


export function resolveRegionMapFeatureSet(labels: string[], geography: Record<string, unknown> | null) {
  const cultureRegion = typeof geography?.cultureRegion === "string"
    ? geography.cultureRegion.trim().toUpperCase()
    : "";
  const countryMatches = labels.filter((label) => resolveRegionMapFeature(label, "country") != null).length;
  const stateMatches = labels.filter((label) => resolveRegionMapFeature(label, "us-state") != null).length;
  if (cultureRegion === "US" && stateMatches > 0 && stateMatches >= countryMatches) {
    return "us-state" as const;
  }
  return "country" as const;
}

export function getRegionMapBaseFeatures(featureSet: "country" | "us-state") {
  return featureSet === "us-state" ? US_STATE_FEATURES : WORLD_COUNTRY_FEATURES;
}

export type RegionMapResources = typeof import("./region-map-geography");
