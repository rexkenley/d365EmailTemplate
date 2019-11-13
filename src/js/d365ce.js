/* global Xrm */
import { get } from "lodash";

const headers = {
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    Accept: "application/json",
    "Content-Type": "application/json; charset=utf-8",
    Prefer: `odata.include-annotations="*"` // eslint-disable-line
  },
  apiVersion = "v9.1";

/**
 * Gets the metadata of the listed entities
 * @param  {...string} entities
 * @return {Promise<Object>}
 */
export async function getMetaData(...entities) {
  try {
    if (!entities || !entities.length) return [];

    const { getClientUrl } = Xrm.Utility.getGlobalContext(),
      metas = {},
      entitySetNames = {};

    // eslint-disable-next-line
    for (const e of entities) {
      // eslint-disable-next-line
      const result = await fetch(
          `${getClientUrl()}/api/data/${apiVersion}/EntityDefinitions(LogicalName='${e}')` +
            "?$select=LogicalName,DisplayName,EntitySetName&$expand=Attributes($select=LogicalName,DisplayName,AttributeType;$filter=IsValidForForm eq true)",
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

/**
 * Retrieves a single entity record
 * @param {string} logicalName
 * @param {string} id
 * @return {Promise<Object>}
 */
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

/**
 * Saves a single entity record
 * @param {string} logicalName
 * @param {string} entity
 * @return {Promise<Object>}
 */
export async function saveEntityData(logicalName, entity) {
  try {
    const { createRecord, updateRecord } = Xrm.WebApi,
      { id, ...x } = entity;
    let result;

    if (entity.id) {
      result = await updateRecord(logicalName, id, { ...x });
      return entity;
    }

    result = await createRecord(logicalName, { ...x });
    return { ...x, id: result.id };
  } catch (e) {
    console.error(e.message || e);
    return null;
  }
}

/**
 * @const {Object}
 */
export const FormatValue = { both: 0, formatOnly: 1, valueOnly: 2 };

/**
 * Formats the object to a more standard format
 * @param {Object} obj
 * @param {number} formatValue
 * @returns {Object}
 */
export function formatObject(obj, formatValue = FormatValue.both) {
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
      switch (formatValue) {
        case FormatValue.formatOnly:
          newObj[k.substring(1).replace("_value", "")] =
            obj[`${k}${formattedTag}`];
          break;

        case FormatValue.valueOnly:
          newObj[k.substring(1).replace("_value", "")] = obj[k];
          break;

        default:
          newObj[k.substring(1).replace("_value", "")] = {
            id: obj[k],
            formatted: obj[`${k}${formattedTag}`],
            logicalName: obj[`${k}${lookupTag}`]
          };
      }
    });

    formatted.forEach(k => {
      switch (formatValue) {
        case FormatValue.formatOnly:
          newObj[k] = obj[`${k}${formattedTag}`];
          break;

        case FormatValue.valueOnly:
          newObj[k] = obj[k];
          break;

        default:
          newObj[k] = {
            value: obj[k],
            formatted: obj[`${k}${formattedTag}`]
          };
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

/**
 * Gets multiple records based on odata
 * @param {string} oData
 * @return {Promise<Object[]>}
 */
export async function getMultipleData(oData) {
  try {
    if (!oData) return null;

    // eslint-disable-next-line
    const { getClientUrl } = Xrm.Utility.getGlobalContext(),
      result = await fetch(
        `${getClientUrl()}/api/data/${apiVersion}/${oData}`,
        {
          method: "GET",
          headers
        }
      ),
      json = result && (await result.json()),
      data = json.value.map(r => formatObject(r));

    return data;
  } catch (e) {
    console.error(e.message || e);
    return null;
  }
}
