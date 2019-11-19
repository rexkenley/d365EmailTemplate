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
import "tinymce/tinymce";
import "tinymce/themes/silver/theme";
import "tinymce/skins/ui/oxide/skin.min.css";
import "tinymce/skins/ui/oxide/content.min.css";
import "tinymce/skins/content/default/content.css";
import "tinymce/plugins/visualchars/index";
import "tinymce/plugins/visualblocks/index";
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
import "tinymce/plugins/autoresize/index";
import "tinymce/plugins/searchreplace/index";
import { Editor as TinyEditor } from "@tinymce/tinymce-react";
import get from "lodash/get";

import { getEntityData, formatObject, saveEntityData } from "../js/d365ce";
import merge from "../js/merge";
import {
  setEntity,
  setTemplate,
  getTemplates,
  setAttribute
} from "../js/store";

/**
 * @param {Object} tinyEditor
 * @param {Object} meta
 * @param {Object[]} templates
 * @param {string} entity
 * @param {Object} template
 * @param {string} attribute
 * @param {Object} regardingObjectId
 */
const getItems = (
    tinyEditor,
    meta,
    templates,
    entity,
    template,
    attribute,
    regardingObjectId
  ) => {
    const dispatch = useDispatch(),
      smpTemplateItems = [
        {
          key: "newTemplate",
          text: "New Template",
          iconProps: { iconName: "WebTemplate" },
          onClick: () => {
            tinyEditor.current.editor.setContent("");
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
            canCheck: true,
            checked: t.subject === get(template, "subject"),
            iconProps: { iconName: "QuickNote" },
            onClick: () => {
              const { annotationid: id, subject, notetext } = t;

              tinyEditor.current.editor.setContent(notetext);

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
          canCheck: true,
          checked: entity === "global",
          onClick: () => {
            dispatch(setEntity("global"));
          }
        },
        ...Object.keys(meta)
          .filter(k => ["annotation", "email"].every(e => e !== k))
          .map(k => ({
            key: k,
            text: meta[k].displayName,
            canCheck: true,
            checked: entity === k,
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
        }
      ],
      mergeHtmlData = async (showPreview = false) => {
        const { id, logicalName } = regardingObjectId,
          data = await getEntityData(logicalName, id),
          html = merge(template.notetext, data);

        if (showPreview) {
          window.open("about:blank", "", "_blank").document.write(html);
        } else {
          window.opener.Xrm.Page.getAttribute("description").setValue(html);
        }
      },
      cbItems2 = [
        {
          key: "save",
          text: "Save",
          iconProps: { iconName: "SaveTemplate" },
          onClick: async () => {
            if (!template || !template.subject) return;

            const updated = {
              ...template,
              notetext: tinyEditor.current.editor.getContent()
            };

            await saveEntityData("annotation", updated);
            dispatch(setTemplate(updated));
            dispatch(getTemplates());
          }
        },
        {
          key: "merge",
          text: "Merge",
          iconProps: { iconName: "Merge" },
          disabled: !regardingObjectId,
          onClick: () => {
            mergeHtmlData();
          }
        },
        {
          key: "preview",
          text: "Preview",
          iconProps: { iconName: "Preview" },
          disabled: !regardingObjectId,
          onClick: () => {
            mergeHtmlData(true);
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
            .sort(
              (a, b) =>
                (a.displayName < b.displayName && -1) ||
                (a.displayName > b.displayName && 1) ||
                0
            )
            .map(a => {
              if (a.attributeType === "DateTime") {
                return {
                  key: a.logicalName,
                  text: a.displayName,
                  canCheck: true,
                  checked: attribute === a.logicalName,
                  onClick: () => {
                    dispatch(setAttribute(a.logicalName));
                  }
                };
              }

              return {
                key: a.logicalName,
                text: a.displayName,
                onClick: () => {
                  if (!template) return;

                  const useFormatted = [
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
                    ].includes(a.attributeType),
                    notetext = `${template.notetext} {{${a.logicalName +
                      (useFormatted ? ".formatted" : "")}}}`;

                  tinyEditor.current.editor.setContent(notetext);

                  dispatch(setAttribute(""));
                  dispatch(
                    setTemplate({
                      ...template,
                      notetext
                    })
                  );
                }
              };
            })
        }
      });
    }

    if (attribute) {
      const isDateTime =
          meta[entity].attributes.find(a => a.logicalName === attribute)
            .attributeType === "DateTime",
        setNoteText = notetext => {
          tinyEditor.current.editor.setContent(notetext);
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
        cbItems.push({
          key: "format",
          text: "Format",
          iconProps: { iconName: "DateTime" },
          subMenuProps: {
            items: [
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
            ]
          }
        });
      }
    }

    return cbItems.concat(cbItems2);
  },
  tinyEditor = React.createRef(),
  Editor = () => {
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
      },
      fpCB = cb => {
        const input = document.createElement("input");
        input.setAttribute("type", "file");
        input.setAttribute("accept", "image/*");

        input.onchange = () => {
          const file = input.files[0],
            reader = new FileReader();

          reader.onload = () => {
            const id = `blobid${new Date().getTime()}`,
              { blobCache } = tinyEditor.current.editor.editorUpload,
              base64 = reader.result.split(",")[1],
              blobInfo = blobCache.create(id, file, base64);

            blobCache.add(blobInfo);

            cb(blobInfo.blobUri(), { title: file.name });
          };
          reader.readAsDataURL(file);
        };

        input.click();
      },
      editorHeight = window.innerHeight - 80;

    return (
      <Fabric>
        <CommandBar
          items={getItems(
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
          init={{
            skin: false,
            content_css: false,
            plugins:
              "autoresize, searchreplace, visualchars, visualblocks, image, imagetools, link, media, codesample, charmap, emoticons, hr, table, help",
            menu: {
              file: {
                title: "File",
                items: "newdocument"
              },
              edit: {
                title: "Edit",
                items: "undo redo | cut copy paste | selectall | searchreplace"
              },
              view: {
                title: "View",
                items: "visualaid visualchars visualblocks"
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
            autoresize_on_init: true,
            autoresize_bottom_margin: 80,
            autoresize_overflow_padding: 50,
            max_height: editorHeight,
            min_height: editorHeight,
            automatic_uploads: true,
            image_advtab: true,
            image_title: true,
            image_description: false,
            file_picker_types: "image",
            file_picker_callback: fpCB,
            toolbar: false
          }}
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
