import "tinymce";
import "tinymce/skins/ui/oxide/skin.min.css";
import "tinymce/skins/ui/oxide/content.min.css";
import "tinymce/skins/content/default/content.css";
import "tinymce/themes/silver/theme";

import "tinymce/plugins/image/index";
import "tinymce/plugins/imagetools/index";
import "tinymce/plugins/link/index";
import "tinymce/plugins/media/index";
import "tinymce/plugins/codesample/index";
import "tinymce/plugins/charmap/index";
import "tinymce/plugins/emoticons/index";
import "tinymce/plugins/emoticons/js/emojis";
import "tinymce/plugins/hr/index";
import "tinymce/plugins/table/index";
import "tinymce/plugins/help/index";

import { get } from "lodash";
import React, { useState } from "react";
import { useSelector, useDispatch } from "react-redux";
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
import { Editor as TinyEditor } from "@tinymce/tinymce-react";

import { getEntityData, formatObject, saveEntityData } from "./../js/d365ce";
import merge from "../js/merge";
import store, {
  setEntity,
  setTemplate,
  getTemplates,
  setAttribute
} from "../js/store";

const getItems = (meta, entity, templates, attribute) => {
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

    if (entity && templates) {
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
            dispatch(setEntity("global"));
          }
        },
        ...Object.keys(meta)
          .filter(k => ["annotation", "email"].every(e => e !== k))
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

            template.notetext = tinyMCE.get()[0].getContent();

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
              data = formatObject(result),
              html = merge(template.notetext, data);

            saveEntityData("email", {
              subject: template.subject,
              description: html
            });
          }
        }
      ];

    if (meta && entity && entity !== "global") {
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
              if (a.attributeType === "DateTime") {
                return {
                  key: a.logicalName,
                  text: a.displayName,
                  onClick: () => {
                    dispatch(setAttribute(a.logicalName));
                  }
                };
              }

              return {
                key: a.logicalName,
                text: a.displayName,
                onClick: () => {
                  const { template } = store.getState(),
                    useFormatted = [
                      "BigInt",
                      "Decimal",
                      "Double",
                      "Integer",
                      "Boolean",
                      "Customer",
                      "Lookup",
                      "Money",
                      "Owner",
                      "Picklist",
                      "State",
                      "Status"
                    ].includes(a.attributeType);

                  if (!template) return;

                  dispatch(setAttribute(""));
                  dispatch(
                    setTemplate({
                      ...template,
                      notetext: `${template.notetext} {{${a.logicalName +
                        (useFormatted ? ".formatted" : "")}}}`
                    })
                  );
                }
              };
            })
        }
      });
    }

    if (attribute) {
      const { template } = store.getState(),
        isDateTime =
          meta[entity].attributes.find(a => a.logicalName === attribute)
            .attributeType === "DateTime",
        setNoteText = notetext => {
          dispatch(setTemplate({ ...template, notetext }));
        },
        defaultItem = {
          key: "default",
          text: "Default",
          onClick: () => {
            setNoteText(`${template.notetext} {{${attribute}.formatted}}`);
          }
        };

      if (isDateTime) {
        const dtItems = [
          defaultItem,
          {
            key: "longDate",
            text: "Long Date",
            onClick: () => {
              setNoteText(
                `${template.notetext} {{formatDate ${attribute}.value month="long" day="numeric" year="numeric"}}`
              );
            }
          },
          {
            key: "shortDate",
            text: "Short Date",
            onClick: () => {
              setNoteText(
                `${template.notetext} {{formatDate ${attribute}.value month="2-digit" day="2-digit" year="2-digit"}}`
              );
            }
          },
          {
            key: "hours12",
            text: "12 Hours",
            onClick: () => {
              setNoteText(
                `${template.notetext} {{formatTime ${attribute}.value hour12=true hour="numeric" minute="numeric"}}`
              );
            }
          },
          {
            key: "hours24",
            text: "24 Hours",
            onClick: () => {
              setNoteText(
                `${template.notetext} {{formatTime ${attribute}.value hour12=false hour="numeric" minute="numeric"}}`
              );
            }
          }
        ];

        cbItems.push({
          key: "format",
          text: "Format",
          iconProps: { iconName: "DateTime" },
          subMenuProps: {
            items: dtItems
          }
        });
      }
    }

    return cbItems;
  },
  Editor = () => {
    const [templateName, setTemplateName] = useState(""),
      dispatch = useDispatch(),
      meta = useSelector(state => state.meta),
      entity = useSelector(state => state.entity),
      template = useSelector(state => state.template),
      templates = useSelector(state => state.templates),
      attribute = useSelector(state => state.attribute),
      dismiss = () => {
        dispatch(setTemplate({}));
        setTemplateName("");
      };

    return (
      <Fabric>
        <CommandBar items={getItems(meta, entity, templates, attribute)} />
        <TinyEditor
          init={{
            height: window.innerHeight - 80,
            width: window.innerWidth - 40,
            plugins:
              "image, imagetools, link, media, codesample, charmap, emoticons, hr, table, help",
            menu: {
              file: {
                title: "File",
                items: "newdocument"
              },
              edit: {
                title: "Edit",
                items: "undo redo | cut copy paste | selectall | searchreplace"
              },
              insert: {
                title: "Insert",
                items: "image link media codesample | charmap emoticons hr"
              },
              format: {
                title: "Format",
                items:
                  "bold italic underline strikethrough superscript subscript codeformat | formats blockformats fontformats fontsizes align | forecolor backcolor | removeformat"
              },
              table: {
                title: "Table",
                items: "inserttable tableprops deletetable row column cell"
              },
              help: { title: "Help", items: "help" }
            },
            image_title: true,
            file_picker_types: "image",
            file_picker_callback: (cb, value, meta) => {
              const input = document.createElement("input");
              input.setAttribute("type", "file");
              input.setAttribute("accept", "image/*");

              input.onchange = function() {
                const file = this.files[0],
                  reader = new FileReader();

                reader.onload = function() {
                  const id = `blobid${new Date().getTime()}`,
                    { blobCache } = tinyMCE.get()[0].editorUpload,
                    base64 = reader.result.split(",")[1],
                    blobInfo = blobCache.create(id, file, base64);

                  blobCache.add(blobInfo);

                  cb(blobInfo.blobUri(), { title: file.name });
                };
                reader.readAsDataURL(file);
              };

              input.click();
            },
            toolbar: false
          }}
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
