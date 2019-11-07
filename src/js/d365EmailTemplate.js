import React from "react";
import ReactDOM from "react-dom";
import { Provider } from "react-redux";
import { initializeIcons } from "@uifabric/icons";
import { getMetaData } from "./d365ce";
import store, { setMeta, getTemplates } from "./store";
import Editor from "../jsx/editor";

initializeIcons();
const run = async () => {
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

  ReactDOM.render(
    <Provider store={store}>
      <Editor />
    </Provider>,
    document.getElementById("d365EmailTemplate")
  );
};

run();
