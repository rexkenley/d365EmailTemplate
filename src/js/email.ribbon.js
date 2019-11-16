/* eslint-disable import/prefer-default-export */
/* global Xrm */
import get from "lodash/get";

export function openD365EmailTemplate(primaryControl) {
  const wrName = "vm_d365EmailTemplate.html";

  if (
    get(
      primaryControl.ui.controls
        .get("regardingobjectid")
        .getAttribute()
        .getValue(),
      "[0]",
      false
    )
  ) {
    const lookup = get(
        primaryControl.ui.controls
          .get("regardingobjectid")
          .getAttribute()
          .getValue(),
        "[0]"
      ),
      { id, typename } = lookup,
      regardingObjectId = {
        id: id.replace(/[{}]/g, ""),
        logicalName: typename
      },
      par = encodeURIComponent(
        `regardingObjectId=${JSON.stringify(regardingObjectId)}`
      );

    Xrm.Navigation.openWebResource(wrName, null, par);
  } else {
    Xrm.Navigation.openWebResource(wrName);
  }
}
