import get from "lodash/get";
import { useDispatch } from "react-redux";

import { getEntityData, saveEntityData } from "./d365ce";
import merge from "./merge";
import { setEntity, setTemplate, getTemplates, setAttribute } from "./store";

/**
 * @module editorCommandBar
 */

/**
 * getCBItems
 * @param {Object} tinyEditor
 * @param {Object} meta
 * @param {Object[]} templates
 * @param {string} entity
 * @param {Object} template
 * @param {string} attribute
 * @param {Object} regardingObjectId
 * @return {Object[]}
 */
const getCBItems = (
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
          await dispatch(getTemplates());
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
          .slice()
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
};

export default getCBItems;
