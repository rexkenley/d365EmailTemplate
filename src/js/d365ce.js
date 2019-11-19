/* global Xrm */
import get from "lodash/get";
import isUUID from "validator/lib/isUUID";

/**
 * @const {Boolean}
 */
export const isV9 = !!Xrm.WebApi;

const { getClientUrl } = isV9
    ? Xrm.Utility.getGlobalContext()
    : Xrm.Page.context,
  apiVersion = isV9 ? "v9.1" : "v8.2",
  headers = {
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    Accept: "application/json",
    "Content-Type": "application/json; charset=utf-8",
    Prefer: `odata.include-annotations="*"` // eslint-disable-line
  };

/**
 * Gets the metadata of the listed entities
 * @param  {...string} entities
 * @return {Promise<Object>}
 */
export async function getMetaData(...entities) {
  try {
    if (!entities || !entities.length) return [];

    const metas = {},
      entitySetNames = {};

    // eslint-disable-next-line
    for (const e of entities) {
      // eslint-disable-next-line
      const result = await fetch(
          `${getClientUrl()}/api/data/${apiVersion}/EntityDefinitions(LogicalName='${e}')` +
            `?$select=LogicalName,DisplayName,EntitySetName&$expand=Attributes($select=LogicalName,DisplayName,AttributeType;${
              isV9 ? "$filter=IsValidForForm eq true" : ""
            })`,
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
 * Retrieves a single entity record
 * @param {string} logicalName
 * @param {string} id
 * @return {Promise<Object>}
 */
export async function getEntityData(logicalName, id) {
  try {
    if (!logicalName || !isUUID(id)) return null;

    let data;

    if (isV9) {
      const result = await Xrm.WebApi.retrieveRecord(
        logicalName,
        id.toLowerCase()
      );

      data = formatObject(result);
    } else {
      const ENTITY_SET_NAMES = JSON.parse(window.ENTITY_SET_NAMES),
        result = await fetch(
          `${getClientUrl()}/api/data/${apiVersion}/${
            ENTITY_SET_NAMES[logicalName]
          }(${id})`,
          {
            method: "GET",
            headers
          }
        ),
        json = result && (await result.json());

      data = formatObject(json);
    }

    return data;
  } catch (e) {
    console.error(e.message || e);
    return null;
  }
}

/**
 * Gets multiple records based on odata
 * @param {string} logicalName
 * @param {string} oData
 * @return {Promise<Object[]>}
 */
export async function getMultipleData(logicalName, oData) {
  try {
    if (!logicalName || !oData) return null;

    let data;

    if (isV9) {
      const result = await Xrm.WebApi.retrieveMultipleRecords(
        logicalName,
        `${oData}`
      );

      data = result.entities.map(r => formatObject(r));
    } else {
      const ENTITY_SET_NAMES = JSON.parse(window.ENTITY_SET_NAMES),
        result = await fetch(
          `${getClientUrl()}/api/data/${apiVersion}/${
            ENTITY_SET_NAMES[logicalName]
          }?${oData}`,
          {
            method: "GET",
            headers
          }
        ),
        json = result && (await result.json());

      data = json.value.map(r => formatObject(r));
    }

    return data;
  } catch (e) {
    console.error(e.message || e);
    return null;
  }
}

/**
 * Saves a single entity record
 * @param {string} logicalName
 * @param {Object} entity
 * @return {Promise<Object>}
 */
export async function saveEntityData(logicalName, entity) {
  try {
    if (!logicalName || !entity) return null;

    const { id, ...x } = entity;
    let result;

    if (isV9) {
      if (isUUID(id)) {
        result = await Xrm.WebApi.updateRecord(logicalName, id, { ...x });
        return entity;
      }

      result = await Xrm.WebApi.createRecord(logicalName, { ...x });
      return { ...x, id: result.id };
    }

    const ENTITY_SET_NAMES = JSON.parse(window.ENTITY_SET_NAMES);

    result = await fetch(
      `${getClientUrl()}/api/data/${apiVersion}/${
        ENTITY_SET_NAMES[logicalName]
      }${isUUID(id) ? `(${id})` : ""}`,
      {
        method: isUUID(id) ? "PATCH" : "POST",
        headers,
        body: JSON.stringify({ ...x })
      }
    );

    if (isUUID(id)) return entity;
    return { ...x, id: result.id };
  } catch (e) {
    console.error(e.message || e);
    return null;
  }
}
