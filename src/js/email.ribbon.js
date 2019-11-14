/* global Xrm */
const openD365EmailTemplate = () => {
  const wrName = "vm_d365EmailTemplate.html",
    regardingObjectId = { id: "", logicalName: "" },
    par = encodeURIComponent(
      `regardingObjectId=${JSON.stringify(regardingObjectId)}`
    ),
    windowOptions = { height: 600, width: 620 };

  Xrm.Utility.getGlobalContext().getQueryStringParameters();

  Xrm.Navigation.openWebResource(wrName, windowOptions, par);
};

export default openD365EmailTemplate;
