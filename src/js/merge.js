import Handlebars from "handlebars";
import HandlebarsIntl from "handlebars-intl";

HandlebarsIntl.registerWith(Handlebars);

export default function merge(source, data) {
  if (!source || !data) return "";

  const template = Handlebars.compile(source);

  return template(data);
}
