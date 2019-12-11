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

import { setTemplate } from "../js/store";
import getCBItems from "../js/editorCommandBar";

import TinyEditor from "./tinyMCEPAC";

/**
 * @module editor
 */

const tinyEditor = React.createRef(),
  Editor = (store, initialValue, onHTMLChange) => {
    const [templateName, setTemplateName] = useState(""),
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
      <Provider store={store}>
        <Fabric>
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
          <TinyEditor
            ref={tinyEditor}
            disabled={!template}
            initialValue={initialValue}
            onHTMLChange={onHTMLChange}
          />
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
        </Fabric>
      </Provider>
    );
  };

export default Editor;

export const setDescription = description => {
  try {
    tinyEditor.current.editor.setContent(description);
  } catch (ex) {
    console.warn(ex.message || ex);
  }
};
