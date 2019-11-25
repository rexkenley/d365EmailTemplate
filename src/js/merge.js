import Handlebars from "handlebars";
import HandlebarsIntl from "handlebars-intl";

HandlebarsIntl.registerWith(Handlebars);

/**
 * @module merge
 */

/**
 * Merge the source with the data and returns an html string
 * @param {string} source
 * @param {Object} data
 * @return {string}
 */
export default function merge(source, data) {
  if (!source || !data) return "";

  const template = Handlebars.compile(source);

  return template(data);
}
