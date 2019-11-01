import { merge } from "../../src/js/merge";

const source = "<div><p>{{message}}</p></div>",
  data = { message: "Test" };

test("Merge Basic Test", () => {
  expect(merge(source, data)).toBe("<div><p>Test</p></div>");
});
