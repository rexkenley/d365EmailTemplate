/* global Xrm */
import { get } from "lodash";
import { is } from "@babel/types";

const { getGlobalContext } = get(window, "Xrm.Utility", {}),
  headers = {
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    Accept: "application/json",
    "Content-Type": "application/json; charset=utf-8",
    Prefer: `odata.include-annotations="*"` // eslint-disable-line
  };

export async function getMetaData(...entities) {
  try {
    if (!entities || !entities.length) return [];

    const { getClientUrl } = getGlobalContext(),
      metas = {},
      entitySetNames = {};

    // eslint-disable-next-line
    for (const e of entities) {
      // eslint-disable-next-line
      const result = await fetch(
          `${getClientUrl()}/api/data/v9.1/EntityDefinitions(LogicalName='${e}')` +
            "?$select=LogicalName,DisplayName,EntitySetName&$expand=Attributes($select=LogicalName,DisplayName,AttributeType;$filter=IsValidForRead eq true)",
          {
            method: "GET",
            headers
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

      entitySetNames[e] = get(json, "EntitySetName", "");
    }

    // https://debajmecrm.com/2018/09/29/cannot-read-property-entity-name-of-null-error-while-executing-a-bound-action-from-from-a-webresource-in-dynamics-v9-0-using-xrm-webapi-execute/
    if (!get(window, "ENTITY_SET_NAMES")) {
      window.ENTITY_SET_NAMES = JSON.stringify(entitySetNames);
    }

    return metas;
  } catch (e) {
    console.error(e.message || e);
    return null;
  }
}

export function formatObject(obj, isFormattedOnly = false) {
  try {
    if (!obj) return null;

    const lookupTag = "@Microsoft.Dynamics.CRM.lookuplogicalname",
      lookups = Object.keys(obj)
        .filter(k => k.endsWith(lookupTag))
        .map(k => k.replace(lookupTag, "")),
      formattedTag = "@OData.Community.Display.V1.FormattedValue",
      formatted = Object.keys(obj)
        .filter(
          k => !lookups.some(l => k.startsWith(l)) && k.endsWith(formattedTag)
        )
        .map(k => k.replace(formattedTag, "")),
      others = Object.keys(obj).filter(
        k =>
          !["@odata.context", "@odata.etag"].includes(k) &&
          !lookups.some(l => k.startsWith(l)) &&
          !formatted.some(f => k.startsWith(f))
      ),
      newObj = {};

    lookups.forEach(k => {
      if (isFormattedOnly) {
        newObj[k.substring(1).replace("_value", "")] =
          obj[`${k}${formattedTag}`];
      } else {
        const newLookUp = {
          id: obj[k],
          formatted: obj[`${k}${formattedTag}`],
          logicalName: obj[`${k}${lookupTag}`]
        };

        newObj[k.substring(1).replace("_value", "")] = newLookUp;
      }
    });

    formatted.forEach(k => {
      if (isFormattedOnly) {
        newObj[k] = obj[`${k}${formattedTag}`];
      } else {
        const newFormatted = {
          value: obj[k],
          formatted: obj[`${k}${formattedTag}`]
        };

        newObj[k] = newFormatted;
      }
    });

    others.forEach(k => {
      newObj[k] = obj[k];
    });

    return newObj;
  } catch (e) {
    console.error(e.message || e);
    return null;
  }
}

export async function getEntityData(logicalName, id) {
  try {
    const result = await Xrm.WebApi.retrieveRecord(
      logicalName,
      id.toLowerCase()
    );

    return result;
  } catch (e) {
    console.error(e.message || e);
    return null;
  }
}
