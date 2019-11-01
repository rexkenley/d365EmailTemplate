import "react-quill/dist/quill.snow.css";
import "../css/fonts.css";

import React, { useState } from "react";
import ReactQuill from "react-quill";

import { Fabric } from "office-ui-fabric-react/lib/Fabric";
import { CommandBar } from "office-ui-fabric-react/lib/CommandBar";

const { Quill } = ReactQuill,
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

const getItems = meta => {
    const subMenuProps = {
        items: Object.keys(meta).map(k => ({
          key: meta[k].metadataId,
          text: meta[k].displayName,
          onClick: () => {
            console.log(k);
          }
        }))
      },
      cbItems = [
        {
          key: "entity",
          text: "Entity",
          iconProps: { iconName: "FileTemplate" },
          subMenuProps
        },
        {
          key: "new",
          text: "New",
          iconProps: { iconName: "FileTemplate" },
          onClick: () => {}
        },
        {
          key: "save",
          text: "Save",
          iconProps: { iconName: "SaveTemplate" },
          onClick: () => {}
        }
      ];

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
  Editor = ({ meta }) => {
    const [text, setText] = useState(""),
      [selectedEntity, setSelectedEntity] = useState("");

    return (
      <Fabric>
        <CommandBar items={getItems(meta)} />
        <ReactQuill
          modules={modules}
          formats={formats}
          value={text}
          onChange={value => setText(value)}
        />
      </Fabric>
    );
  };

export default Editor;
