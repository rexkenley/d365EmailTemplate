import { configureStore, createSlice } from "@reduxjs/toolkit";
import { getMultipleData } from "./d365ce";

const editorSlice = createSlice({
    name: "editor",
    initialState: {
      meta: null,
      entity: "",
      template: null,
      templates: null,
      attribute: "",
      regardingObjectId: null
    },
    reducers: {
      setMeta(state, { payload }) {
        return { ...state, meta: payload };
      },
      setEntity(state, { payload }) {
        return { ...state, entity: payload };
      },
      setTemplate(state, { payload }) {
        return { ...state, template: payload };
      },
      setTemplates(state, { payload }) {
        return { ...state, templates: payload };
      },
      setAttribute(state, { payload }) {
        return { ...state, attribute: payload };
      },
      setRegardingObjectId(state, { payload }) {
        return {
          ...state,
          regardingObjectId: payload
        };
      }
    }
  }),
  store = configureStore({ reducer: editorSlice.reducer });

export default store;
export const {
  setMeta,
  setEntity,
  setTemplate,
  setTemplates,
  setAttribute,
  setRegardingObjectId
} = editorSlice.actions;

/**
 * Gets all of the Templates from Annotations
 */
export function getTemplates() {
  return async dispatch => {
    const templates = await getMultipleData(
      "annotation",
      "$select=annotationid,subject,notetext&$filter=startswith(subject,'d365EmailTemplate')&$orderby=subject"
    );

    dispatch(setTemplates(templates));
  };
}
