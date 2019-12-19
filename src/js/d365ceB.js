/* global Xrm */
import get from "lodash/get";
import isUUID from "validator/lib/isUUID";

export const getEnvironment = context => {
  const api = get(context, "webAPI", get(context, "WebApi.online")),
    version = api ? "v9.1" : "v8.2",
    headers = {
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      Accept: "application/json",
      "Content-Type": "application/json; charset=utf-8",
      Prefer: `odata.include-annotations="*"` // eslint-disable-line
    };

  return {
    api,
    version,
    headers,
    isV9: !!api
  };
};

export async function getMetaData(context, ...entities) {
  try {
    if (!entities || !entities.length) return [];

    const env = getEnvironment(context),
      metas = {};

    // eslint-disable-next-line
    for (const e of entities) {
      // eslint-disable-next-line
      const result = await fetch(
          `${getClientUrl()}/api/data/${
            env.version
          }/EntityDefinitions(LogicalName='${e}')` +
            `?$select=LogicalName,DisplayName,EntitySetName&$expand=Attributes($select=LogicalName,DisplayName,AttributeType;${
              env.isV9 ? "$filter=IsValidForForm eq true" : ""
            })`,
          {
            method: "GET",
            headers: env.headers
          }
        ),
        json = result && (await result.json()); // eslint-disable-line

      metas[e] = {
        metadataId: get(json, "MetadataId", ""),
        displayName: get(json, "DisplayName.LocalizedLabels[0].Label", ""),
        entitySetName: get(json, "EntitySetName", ""),
        attributes: get(json, "Attributes", [])
          .filter(a => !!get(a, "DisplayName.LocalizedLabels[0].Label"))
          .map(a => ({
            metadataId: get(a, "MetadataId", ""),
            displayName: get(a, "DisplayName.LocalizedLabels[0].Label", ""),
            logicalName: get(a, "LogicalName", ""),
            attributeType: get(a, "AttributeType", "")
          }))
      };
    }

    return metas;
  } catch (e) {
    console.error(e.message || e);
    return null;
  }
}
