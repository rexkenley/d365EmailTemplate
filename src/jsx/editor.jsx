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
import { setTemplate } from "../js/store";

import getCBItems from "../js/editorCommandBar";

/**
 * @module editor
 */

const tinyEditor = React.createRef(),
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
  Editor = () => {
    const [templateName, setTemplateName] = useState(""),
      dispatch = useDispatch(),
      meta = useSelector(state => state.meta),
      entity = useSelector(state => state.entity),
      template = useSelector(state => state.template),
      templates = useSelector(state => state.templates),
      attribute = useSelector(state => state.attribute),
      regardingObjectId = useSelector(state => state.regardingObjectId),
      editorHeight = window.innerHeight - 80,
      dismiss = () => {
        dispatch(setTemplate({}));
        setTemplateName("");
      };

    return (
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
