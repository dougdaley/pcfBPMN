import { IInputs, IOutputs } from "./generated/ManifestTypes";

/**
 * @title pcfBPMN
 * @description A control that allows users to visualise, create and edit BPMN diagrams stored in a Dataverse column.
 * @remarks This control is a wrapper around the bpmn-js library, which is a JavaScript library for viewing and editing BPMN diagrams.
 */

// Import the necessary libraries and modeler diagram styles

import Modeler from 'bpmn-js/lib/Modeler';
import 'bpmn-js/dist/assets/diagram-js.css';
import 'bpmn-js/dist/assets/bpmn-font/css/bpmn.css';
import { EventBus } from "bpmn-js/lib/BaseViewer";
//import { defaultMaxListeners } from "events";
//import { error } from "console";

export class pcfBPMN implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    
    /**
     * Declare the properties of the class
     */

    private _modeler: Modeler; // bpmn-js modeler instance for BPMN diagram rendering and editing

    private _container: HTMLDivElement; // HTML container element into which the modeler rendered

    private _notifyOutputChanged: () => void; // Callback function to notify the framework of output changes

    private _bpmnXML: string; // BPMN XML string that maintains the current diagram and is stored in the Dataverse column

    private _context: ComponentFramework.Context<IInputs>; // Context object that provides access to the framework and control properties
    
    /**
     * Empty constructor.
     */
    constructor() {
        // Empty
    }

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
    ): void {
        this._context = context;
        this._container = container;
        this._container.style.height = "600px"; // Todo: make this dynamic and resizeable by a user
        this._notifyOutputChanged = notifyOutputChanged;

        // Instantiate the BPMN modeler and set the container for rendering
        this._modeler = new Modeler({
            container: this._container
        });

        // Initialise with a default diagram in case no valid XML is able to be loaded from the field
        this._modeler.createDiagram();
        
        // Try to load the diagram from the field context
        if (this._context.parameters.bpmnXML.raw) {
            this._bpmnXML = this._context.parameters.bpmnXML.raw;

            try {
                this._modeler.importXML(this._bpmnXML);
            } catch (error) {
                console.error('Error importing BPMN XML:', error);
            }
        }

        // Listen for changes to the BPMN diagram, update the bpmnXML property and notify the control framework accordingly
        this.attachEventListeners();
    }

    /**
     * Attaches event listeners to the BPMN modeler instance.
     * This method listens for changes in the BPMN diagram and updates the bpmnXML property accordingly.
     */
    private attachEventListeners(): void {
        const eventBus = this._modeler.get<EventBus<Event>>('eventBus');

        eventBus.on('element.changed', async () => {
            try {
                const result = await this._modeler.saveXML({ format: true });
                if (result.xml) {
                    this._bpmnXML = result.xml;
                    this._notifyOutputChanged();
                };
            } catch (error) {
                console.error('Error saving BPMN XML:', error);
            }
        });
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public async updateView(context: ComponentFramework.Context<IInputs>): Promise<void> {
        try {
            await this._modeler.importXML(this._bpmnXML);
        } catch (error) {
            console.error('Error importing BPMN XML:', error);
        }
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        return {bpmnXML: this._bpmnXML ? this._bpmnXML : undefined};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }
}
