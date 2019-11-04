import React from "react";
import ReactDOM from "react-dom";
import { initializeIcons } from "@uifabric/icons";
import { getMetaData, getEntityData } from "./d365ce";
import Editor from "../jsx/editor";

getMetaData(
  "account",
  "contact",
  "incident",
  "invoice",
  "lead",
  "opportunity",
  "quote",
  "salesorder"
).then(meta => {
  initializeIcons();
  ReactDOM.render(
    <Editor meta={meta} />,
    document.getElementById("d365EmailTemplate")
  );

  getEntityData("accounts", "3CA3B8D2-034B-E911-A82F-000D3A17CE77");
});
