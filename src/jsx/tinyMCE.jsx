import React from "react";
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
import { Editor } from "@tinymce/tinymce-react";

/**
 * @module tinyMCE
 */

const TinyEditor = React.forwardRef((props, ref) => {
  const { disabled } = props,
    editorHeight = window.innerHeight - 80,
    fpCB = cb => {
      const input = document.createElement("input");
      input.setAttribute("type", "file");
      input.setAttribute("accept", "image/*");

      input.onchange = () => {
        const file = input.files[0],
          reader = new FileReader();

        reader.onload = () => {
          const id = `blobid${new Date().getTime()}`,
            { blobCache } = ref.current.editor.editorUpload,
            base64 = reader.result.split(",")[1],
            blobInfo = blobCache.create(id, file, base64);

          blobCache.add(blobInfo);

          cb(blobInfo.blobUri(), { title: file.name });
        };
        reader.readAsDataURL(file);
      };

      input.click();
    };

  return (
    <Editor
      ref={ref}
      disabled={disabled}
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
  );
});

export default TinyEditor;
