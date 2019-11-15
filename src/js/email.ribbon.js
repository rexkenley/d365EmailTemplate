/* eslint-disable import/prefer-default-export */
/* global Xrm */
export function openD365EmailTemplate() {
  const wrName = "vm_d365EmailTemplate.html",
    regardingObjectId = {
      logicalName: "accounts",
      id: "3CA3B8D2-034B-E911-A82F-000D3A17CE77"
    },
    par = encodeURIComponent(
      `regardingObjectId=${JSON.stringify(regardingObjectId)}`
    ),
    windowOptions = { height: 600, width: 620 };

  Xrm.Navigation.openWebResource(wrName, windowOptions, par);
}
