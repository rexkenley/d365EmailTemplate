import { createStore, applyMiddleware } from "redux";
import thunk from "redux-thunk";
import { getMultipleData } from "./d365ce";

/**
 * @param {Object} state
 * @param {Object} param1
 * @return {Object}
 */
function reducer(state, { type = "", payload = null }) {
  switch (type) {
    case "SET_META":
      return { ...state, meta: payload };
    case "SET_ENTITY":
      return { ...state, entity: payload };
    case "SET_TEMPLATE":
      return { ...state, template: payload };
    case "SET_TEMPLATES":
      return { ...state, templates: payload };
    case "SET_ATTRIBUTE":
      return { ...state, attribute: payload };
    case "SET_REGARDINGOBJECTID":
      return { ...state, regardingObjectId: payload };
    default:
      return state;
  }
}

/**
 * Sets the Meta Object
 * @param {Object} meta
 */
export function setMeta(meta) {
  return dispatch => {
    dispatch({ type: "SET_META", payload: meta });
  };
}

/**
 * Sets the Entity String
 * @param {string} entity
 */
export function setEntity(entity) {
  return dispatch => {
    dispatch({ type: "SET_ENTITY", payload: entity });
  };
}

/**
 * Sets the current Template Object
 * @param {Object} template
 */

export function setTemplate(template) {
  return dispatch => {
    dispatch({ type: "SET_TEMPLATE", payload: template });
  };
}

/**
 * Sets the Templates Array
 * @param {Object[]} templates
 */
export function setTemplates(templates) {
  return dispatch => {
    dispatch({ type: "SET_TEMPLATES", payload: templates });
  };
}

/**
 * Gets all of the Templates from Annotations
 */
export function getTemplates() {
  return async dispatch => {
    const templates = await getMultipleData(
      "annotation",
      "$select=annotationid,subject,notetext&$filter=startswith(subject,'d365EmailTemplate')&$orderby=subject"
    );
    dispatch({ type: "SET_TEMPLATES", payload: templates });
  };
}

/**
 * Sets the Attribute String
 * @param {string} attribute
 */
export function setAttribute(attribute) {
  return dispatch => {
    dispatch({ type: "SET_ATTRIBUTE", payload: attribute });
  };
}

/**
 * Sets the RegardingObjectId
 * @param {Object} regardingObjectId
 */
export function setRegardingObjectId(regardingObjectId) {
  return dispatch => {
    dispatch({ type: "SET_REGARDINGOBJECTID", payload: regardingObjectId });
  };
}

/**
 * @const {Object}
 */
const initialState = {
    meta: null,
    entity: "",
    template: null,
    templates: null,
    attribute: "",
    regardingObjectId: null
  },
  store = createStore(reducer, initialState, applyMiddleware(thunk));

export default store;
