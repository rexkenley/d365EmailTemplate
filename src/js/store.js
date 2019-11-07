import { createStore, applyMiddleware } from "redux";
import thunk from "redux-thunk";
import { getMultipleData } from "./d365ce";

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

    default:
      return state;
  }
}

export function setMeta(meta) {
  return dispatch => {
    dispatch({ type: "SET_META", payload: meta });
  };
}

export function setEntity(entity) {
  return dispatch => {
    dispatch({ type: "SET_ENTITY", payload: entity });
  };
}

export function setTemplate(template) {
  return dispatch => {
    dispatch({ type: "SET_TEMPLATE", payload: template });
  };
}

export function setTemplates(templates) {
  return dispatch => {
    dispatch({ type: "SET_TEMPLATES", payload: templates });
  };
}

export function getTemplates() {
  return async dispatch => {
    const templates = await getMultipleData(
      "annotations?$select=annotationid,subject,notetext&$filter=startswith(subject,'d365EmailTemplate')&$orderby=subject"
    );
    dispatch({ type: "SET_TEMPLATES", payload: templates });
  };
}

export function setAttribute(attribute) {
  return dispatch => {
    dispatch({ type: "SET_ATTRIBUTE", payload: attribute });
  };
}

const initialState = {
    meta: null,
    entity: "",
    template: null,
    templates: null,
    attribute: ""
  },
  store = createStore(reducer, initialState, applyMiddleware(thunk));

export default store;
