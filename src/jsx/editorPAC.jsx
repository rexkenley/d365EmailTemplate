/* global Xrm */
import React, { useState } from "react";
import { Provider, useSelector, useDispatch } from "react-redux";
import { Fabric } from "office-ui-fabric-react/lib/Fabric";
import { CommandBar } from "office-ui-fabric-react/lib/CommandBar";
import {
  Dialog,
  DialogType,
  DialogFooter
} from "office-ui-fabric-react/lib/Dialog";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import {
  PrimaryButton,
  DefaultButton
} from "office-ui-fabric-react/lib/Button";
import { Modal } from "office-ui-fabric-react/lib/Modal";
import get from "lodash/get";

import { setTemplate } from "../js/store";
import getCBItems from "../js/editorCommandBar";

import TinyEditor from "./tinyMCEPAC";

/**
 * @module editor
 */

const tinyEditor = React.createRef(),
  tinyTemplate = React.createRef(),
  Editor = (initialValue, onHTMLChange) => {
    const [showTemplate, setShowTemplate] = useState(false),
      [templateName, setTemplateName] = useState(""),
      dispatch = useDispatch(),
      meta = useSelector(state => state.meta),
      entity = useSelector(state => state.entity),
      template = useSelector(state => state.template),
      templates = useSelector(state => state.templates),
      attribute = useSelector(state => state.attribute),
      regardingObjectId = useSelector(state => state.regardingObjectId),
      dismiss = () => {
        dispatch(setTemplate({}));
        setTemplateName("");
      };

    return (
      <Fabric>
        <TinyEditor
          ref={tinyEditor}
          initialValue={initialValue}
          onHTMLChange={onHTMLChange}
          onTemplatesAction={() => {
            setShowTemplate(true);
          }}
        />
        <Modal
          isOpen={showTemplate}
          onDismiss={() => {
            setShowTemplate(false);
          }}
          isModeless={true}
        >
          <CommandBar
            items={getCBItems(
              tinyEditor,
              meta,
              templates,
              entity,
              template,
              attribute,
              regardingObjectId
            )}
          />
          <TinyEditor ref={tinyTemplate} disabled={!template} />
          <Dialog
            hidden={!entity || !template || template.subject}
            dialogContentProps={{
              type: DialogType.normal,
              title: "New Template"
            }}
            onDismiss={dismiss}
          >
            <TextField
              placeholder="Please enter template name"
              onChange={(ev, value) => {
                setTemplateName(value);
              }}
            />
            <DialogFooter>
              <PrimaryButton
                onClick={() => {
                  dispatch(
                    setTemplate({
                      id: "",
                      subject: `d365EmailTemplate:${entity}:${templateName}`,
                      notetext: ""
                    })
                  );

                  setTemplateName("");
                }}
                text="Ok"
              />
              <DefaultButton onClick={dismiss} text="Cancel" />
            </DialogFooter>
          </Dialog>
        </Modal>
      </Fabric>
    );
  },
  EditorPAC = (store, initialValue, onHTMLChange) => {
    return (
      <Provider store={store}>
        <Editor initialValue={initialValue} onHTMLChange={onHTMLChange} />
      </Provider>
    );
  };

export default EditorPAC;

export const setDescription = description => {
  try {
    tinyEditor.current.editor.setContent(description);
  } catch (ex) {
    console.warn(ex.message || ex);
  }
};

/**
 * Gets the metadata of the listed entities
 * @param  {...string} entities - list of entity names
 * @return {Promise<MetaEntity[]>}
 */
export async function getMetaData(context, ...entities) {
  try {
    if (!entities || !entities.length) return [];

    console.log("getMetaData");

    const globalContext = Xrm.Utility.getGlobalContext(),
      apiVersion = "9.1",
      headers = {
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        Accept: "application/json",
        "Content-Type": "application/json; charset=utf-8",
        Prefer: `odata.include-annotations="*"` // eslint-disable-line
      },
      metas = {};

    // eslint-disable-next-line
    for (const e of entities) {
      // eslint-disable-next-line
      const result = await fetch(
          `${globalContext.getClientUrl()}/api/data/${apiVersion}/EntityDefinitions(LogicalName='${e}')` +
            `?$select=LogicalName,DisplayName,EntitySetName&$expand=Attributes($select=LogicalName,DisplayName,AttributeType;$filter=IsValidForForm eq true)`,
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
    }

    console.log(JSON.stringify(metas));
    return metas;
  } catch (e) {
    console.error(e.message || e);
    return null;
  }
}
