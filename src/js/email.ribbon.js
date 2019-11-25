/* eslint-disable indent */
/* eslint-disable import/prefer-default-export */
/* global Xrm */
import get from "lodash/get";

import { isV9 } from "./d365ce";

/**
 * @module email.ribbon
 */

/**
 * Opens the d365 Email Template html
 * @param {Object} primaryControl
 */
export function openD365EmailTemplate(primaryControl) {
  try {
    const wrName = "vm_d365EmailTemplate.html",
      lookup = isV9
        ? get(
            primaryControl.ui.controls
              .get("regardingobjectid")
              .getAttribute()
              .getValue(),
            "[0]",
            false
          )
        : get(
            Xrm.Page.getAttribute("regardingobjectid").getValue(),
            "[0]",
            false
          );

    if (lookup) {
      const regardingObjectId = {
          id: lookup.id.replace(/[{}]/g, ""),
          logicalName: lookup.typename || lookup.entityType
        },
        par = encodeURIComponent(
          `regardingObjectId=${JSON.stringify(regardingObjectId)}`
        );

      if (isV9) {
        Xrm.Navigation.openWebResource(wrName, null, par);
      } else {
        Xrm.Utility.openWebResource(wrName, par);
      }
    } else if (isV9) {
      Xrm.Navigation.openWebResource(wrName);
    } else {
      Xrm.Utility.openWebResource(wrName);
    }
  } catch (e) {
    console.error(e.message || e);
  }
}
