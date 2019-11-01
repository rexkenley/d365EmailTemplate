import { merge } from "../../src/js/merge";

test("Merge Basic Test", () => {
  const source = "<div><p>{{test}}</p></div>",
    data = { test: "Test" };

  expect(merge(source, data)).toBe("<div><p>Test</p></div>");
});

test("Merge Intl Date Test", () => {
  const source = `<div><p>{{formatDate test day="numeric" month="long" year="numeric" }}</p></div>`,
    data = { test: "1/1/2019" };

  expect(merge(source, data)).toBe("<div><p>January 1, 2019</p></div>");
});

test("Merge Intl Number Test", () => {
  const source = `<ul><li>{{formatNumber test}}</li><li>{{formatNumber test style="percent"}}</li><li>{{formatNumber test style="currency" currency="USD"}}</li></ul>`,
    data = { test: 10 };

  expect(merge(source, data)).toBe(
    "<ul><li>10</li><li>1,000%</li><li>$10.00</li></ul>"
  );
});
