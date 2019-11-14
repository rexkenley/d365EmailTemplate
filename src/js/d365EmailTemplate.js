import React from "react";
import ReactDOM from "react-dom";
import { Provider } from "react-redux";
import { initializeIcons } from "@uifabric/icons";
import { getMetaData } from "./d365ce";
import store, {
  setMeta,
  getTemplates,
  setEntity,
  setRegardingObjectId
} from "./store";
import Editor from "../jsx/editor";

initializeIcons();
const regardingObjectId = {
    logicalName: "accounts",
    id: "3CA3B8D2-034B-E911-A82F-000D3A17CE77"
  },
  run = async () => {
    const meta = await getMetaData(
      "account",
      "contact",
      "incident",
      "invoice",
      "lead",
      "opportunity",
      "quote",
      "salesorder",
      "annotation",
      "email"
    );

    store.dispatch(setMeta(meta));
    store.dispatch(getTemplates());

    if (regardingObjectId) {
      store.dispatch(setRegardingObjectId(regardingObjectId));

      if (Object.keys(meta).some(k => k === regardingObjectId.logicalName)) {
        store.dispatch(setEntity(regardingObjectId.logicalName));
      } else {
        store.dispatch(setEntity("global"));
      }
    } else {
      store.dispatch(setEntity("global"));
    }

    ReactDOM.render(
      <Provider store={store}>
        <Editor />
      </Provider>,
      document.getElementById("d365EmailTemplate")
    );
  };

run();
