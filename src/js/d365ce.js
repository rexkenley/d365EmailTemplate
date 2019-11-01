/* global Xrm */
import { get } from "lodash";

const { getGlobalContext } = Xrm.Utility,
  headers = {
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    Accept: "application/json",
    "Content-Type": "application/json; charset=utf-8",
    Prefer: `odata.include-annotations="*"` // eslint-disable-line
  };

export async function getMetaData(...entities) {
  if (!entities || !entities.length) return [];

  const { getClientUrl } = getGlobalContext(),
    metas = {};

  // eslint-disable-next-line
  for (const e of entities) {
    // eslint-disable-next-line
    const result = await fetch(
        `${getClientUrl()}/api/data/v9.1/EntityDefinitions(LogicalName='${e}')` +
          "?$select=LogicalName,DisplayName&$expand=Attributes($select=LogicalName,DisplayName;$filter=IsValidForRead eq true)",
        {
          method: "GET",
          headers
        }
      ),
      json = result && (await result.json()); // eslint-disable-line

    metas[e] = {
      metadataId: get(json, "MetadataId", ""),
      displayName: get(json, "DisplayName.LocalizedLabels[0].Label", ""),
      attributes: get(json, "Attributes", [])
        .filter(a => !!get(a, "DisplayName.LocalizedLabels[0].Label"))
        .map(a => ({
          metadataId: get(a, "MetadataId", ""),
          displayName: get(a, "DisplayName.LocalizedLabels[0].Label", ""),
          logicalName: get(a, "LogicalName", "")
        }))
    };
  }

  return metas;
}

export function t() {}
