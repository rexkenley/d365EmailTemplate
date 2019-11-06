import store, { setMeta, setEntity, setTemplate } from "../../src/js/store";

test("Reducer SET_META Test", () => {
  store.dispatch(setMeta({}));

  const { meta } = store.getState();
  expect(meta).not.toBeNull();
});

test("Reducer SET_ENTITY Test", () => {
  store.dispatch(setEntity("entity"));

  const { entity } = store.getState();
  expect(entity).toBe("entity");
});

test("Reducer SET_TEMPLATE Test", () => {
  store.dispatch(setTemplate({}));

  const { template } = store.getState();
  expect(template).not.toBeNull();
});
