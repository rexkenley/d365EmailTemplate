import "react-quill/dist/quill.snow.css";
import "../css/fonts.css";

import { get } from "lodash";
import React, { useState } from "react";
import { useSelector, useDispatch } from "react-redux";
import ReactQuill from "react-quill";
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

import {
  getEntityData,
  formatObject,
  saveEntityData,
  FormatValue
} from "./../js/d365ce";
import merge from "../js/merge";
import store, { setEntity, setTemplate, getTemplates } from "../js/store";

const editorRef = React.createRef(),
  { Quill } = ReactQuill,
  Font = Quill.import("formats/font"),
  Size = Quill.import("formats/size");

Font.whitelist = [
  "arial",
  "comic-sans",
  "courier-new",
  "georgia",
  "helvetica",
  "lucida"
  /*"arial-black",
  "tahoma",
  "verdana",
  "garamond",
  "times-new-roman",
  "ms-gothic"*/
];
Quill.register(Font, true);

Size.whitelist = ["extra-small", "small", "medium", "large"];
Quill.register(Size, true);

const getItems = (meta, entity, templates) => {
    const dispatch = useDispatch(),
      smpTemplateItems = [
        {
          key: "newTemplate",
          text: "New Template",
          iconProps: { iconName: "WebTemplate" },
          onClick: () => {
            dispatch(setTemplate({ subject: "" }));
          }
        }
      ];

    if (entity) {
      templates
        .filter(t => t.subject.startsWith(`d365EmailTemplate:${entity}:`))
        .forEach(t => {
          smpTemplateItems.push({
            key: t.subject,
            text: t.subject.replace(`d365EmailTemplate:${entity}:`, ""),
            iconProps: { iconName: "QuickNote" },
            onClick: () => {
              const { annotationid: id, subject, notetext } = t;

              dispatch(
                setTemplate({
                  id,
                  subject,
                  notetext
                })
              );
            }
          });
        });
    }

    const entityItems = [
        {
          key: "global",
          text: "Global",
          onClick: () => {
            dispatch(setEntity(""));
          }
        },
        ...Object.keys(meta)
          .filter(k => k !== "annotation" && k !== "email")
          .map(k => ({
            key: k,
            text: meta[k].displayName,
            onClick: () => {
              dispatch(setEntity(k));
            }
          }))
      ],
      smpEntity = {
        items: entityItems
      },
      smpTemplate = {
        items: smpTemplateItems
      },
      cbItems = [
        {
          key: "entity",
          text: "Entity",
          iconProps: { iconName: "Product" },
          subMenuProps: smpEntity
        },
        {
          key: "templates",
          text: "Templates",
          iconProps: { iconName: "FileTemplate" },
          subMenuProps: smpTemplate
        },
        {
          key: "save",
          text: "Save",
          iconProps: { iconName: "SaveTemplate" },
          onClick: async () => {
            const { template } = store.getState();

            if (!template || !template.subject) return;

            template.notetext = editorRef.current.getEditorContents();

            await saveEntityData("annotation", template);
            dispatch(setTemplate(template));
            dispatch(getTemplates());
          }
        },
        {
          key: "merge",
          text: "Merge",
          iconProps: { iconName: "Merge" },
          onClick: async () => {
            const { template } = store.getState(),
              result = await getEntityData(
                "accounts",
                "3CA3B8D2-034B-E911-A82F-000D3A17CE77"
              ),
              data = formatObject(result, FormatValue.formatOnly),
              html = merge(template.notetext, data);

            saveEntityData("email", {
              subject: template.subject,
              description: html
            });
          }
        }
      ];

    if (meta && entity) {
      cbItems.push({
        key: "attribute",
        text: "Attribute",
        iconProps: { iconName: "ProductList" },
        subMenuProps: {
          items: meta[entity].attributes
            .sort((a, b) => {
              return (
                (a.displayName < b.displayName && -1) ||
                (a.displayName > b.displayName && 1) ||
                0
              );
            })
            .map(a => {
              switch (a.attributeType) {
                case "DateTime":
                  return {
                    key: a.logicalName,
                    text: a.displayName,
                    onClick: () => {}
                  };
                case "BigInt":
                case "Money":
                case "Integer":
                case "Double":
                case "Decimal":
                  return {
                    key: a.logicalName,
                    text: a.displayName,
                    onClick: () => {}
                  };
                default:
                  return {
                    key: a.logicalName,
                    text: a.displayName,
                    onClick: () => {
                      const { template } = store.getState();

                      if (!template) return;

                      dispatch(
                        setTemplate({
                          ...template,
                          notetext: `${template.notetext} {{${a.logicalName}}}`
                        })
                      );
                    }
                  };
              }
            })
        }
      });
    }

    return cbItems;
  },
  modules = {
    toolbar: [
      [{ header: "1" }, { header: "2" }, { font: Font.whitelist }],
      [{ size: Size.whitelist }],
      ["bold", "italic", "underline", "strike", "blockquote"],
      [
        { list: "ordered" },
        { list: "bullet" },
        { indent: "-1" },
        { indent: "+1" }
      ],
      ["link", "image", "video"],
      ["clean"]
    ]
  },
  formats = [
    "header",
    "font",
    "size",
    "bold",
    "italic",
    "underline",
    "strike",
    "blockquote",
    "list",
    "bullet",
    "indent",
    "link",
    "image",
    "video"
  ],
  Editor = () => {
    const [templateName, setTemplateName] = useState(""),
      dispatch = useDispatch(),
      meta = useSelector(state => state.meta),
      template = useSelector(state => state.template),
      templates = useSelector(state => state.templates),
      entity = useSelector(state => state.entity),
      dismiss = () => {
        dispatch(setTemplate({}));
        setTemplateName("");
      };

    return (
      <Fabric>
        <CommandBar items={getItems(meta, entity, templates)} />
        <ReactQuill
          ref={editorRef}
          modules={modules}
          formats={formats}
          value={get(template, "notetext", "")}
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
    );
  };

export default Editor;
