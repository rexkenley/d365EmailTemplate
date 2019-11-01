import React from "react";
import ReactDOM from "react-dom";
import { initializeIcons } from "@uifabric/icons";
import { getMetaData } from "./d365ce";
import Editor from "../jsx/editor";

getMetaData(
  "lead",
  "opportunity",
  "account",
  "contact",
  "quote",
  "salesorder",
  "invoice",
  "incident",
  "contact"
).then(meta => {
  initializeIcons();
  ReactDOM.render(
    <Editor meta={meta} />,
    document.getElementById("d365EmailTemplate")
  );
});
