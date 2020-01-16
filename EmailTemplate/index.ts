import * as React from "react";
import * as ReactDOM from "react-dom";
import { initializeIcons } from "@uifabric/icons";
import { IInputs, IOutputs } from "./generated/ManifestTypes";

import store, {
  setMeta,
  getTemplates,
  setEntity,
  setRegardingObjectId
} from "../src/js/store";
import ET, { setDescription } from "../src/jsx/editorPAC";

export class EmailTemplate
  implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private container: HTMLDivElement;
  private notifyOutputChanged: () => void;
  private current: string;
  private updatedByReact: boolean;

  /**
   * Empty constructor.
   */
  constructor() {}

  /**
   * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
   * Data-set values are not initialized here, use updateView.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
   * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
   * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
   * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
   */
  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ) {
    const { description } = context.parameters,
      initialValue = (description && description.raw) || "",
      run = async () => {
        /*
		https://docs.microsoft.com/en-us/powerapps/developer/component-framework/reference/context
		https://docs.microsoft.com/en-us/powerapps/developer/component-framework/reference/client
		https://docs.microsoft.com/en-us/powerapps/developer/component-framework/reference/entityrecord
		https://docs.microsoft.com/en-us/powerapps/developer/component-framework/reference/navigation
		https://docs.microsoft.com/en-us/powerapps/developer/component-framework/reference/resources
		https://docs.microsoft.com/en-us/powerapps/developer/component-framework/reference/utility
		https://docs.microsoft.com/en-us/powerapps/developer/component-framework/reference/webapi
		

        const meta = await getMetaData(
          "account",
          "contact",
          "incident",
          "invoice",
          "lead",
          "opportunity",
          "quote",
          "salesorder",
          "annotation",
          "email"
        );

        store.dispatch(setMeta(meta));
        await store.dispatch(getTemplates());

        if (regardingObjectId) {
          store.dispatch(setRegardingObjectId(regardingObjectId));

          if (
            Object.keys(meta).some(k => k === regardingObjectId.logicalName)
          ) {
            store.dispatch(setEntity(regardingObjectId.logicalName));
          } else {
            store.dispatch(setEntity("global"));
          }
        } else {
          store.dispatch(setEntity("global"));
		}
		*/

        // Add control initialization code
        ReactDOM.render(
          // @ts-ignore
          React.createElement(
            ET,

            {
              // @ts-ignore
              store,
              initialValue,
              onHTMLChange: (content, editor) => {
                this.current = content;
                this.updatedByReact = true;
                this.notifyOutputChanged();
              }
            }
          ),
          this.container
        );
      };

    this.container = container;
    this.notifyOutputChanged = notifyOutputChanged;
    this.current = initialValue;
    this.updatedByReact = false;

    initializeIcons();
    run();
  }

  /**
   * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
   */
  public updateView(context: ComponentFramework.Context<IInputs>): void {
    // Add code to update control view
    const { description } = context.parameters;

    if (this.updatedByReact) {
      if (this.current === description.raw) this.updatedByReact = false;

      return;
    }

    this.current = description.raw || "";
    setDescription(this.current);
  }

  /**
   * It is called by the framework prior to a control receiving new data.
   * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
   */
  public getOutputs(): IOutputs {
    return { description: this.current };
  }

  /**
   * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
   * i.e. cancelling any pending remote calls, removing listeners, etc.
   */
  public destroy(): void {
    // Add code to cleanup control if necessary
    ReactDOM.unmountComponentAtNode(this.container);
  }
}
