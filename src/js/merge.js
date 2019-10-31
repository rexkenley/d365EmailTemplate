import Handlebars from "handlebars";

let _template;

export function merge(source, data) {
  _template = Handlebars.compile(source);

  return _template(data);
}
