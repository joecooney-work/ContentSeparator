import { IInputs, IOutputs } from "./generated/ManifestTypes";

/**
 * Author: Joe Cooney
 * Company: Microsoft
 * Date: 12.04.2024
 */
export class ContentSeparator implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    //#region Vars
    //Context Logic
    private container: HTMLDivElement;
    private context: ComponentFramework.Context<IInputs>;
    private notifyOutPutChanged: () => void;//to notify form of control change...    
    private state: ComponentFramework.Dictionary;
    
    //Global Vars 
    private cS_Container : HTMLDivElement;
    private cS_Label: HTMLLabelElement;
    private cS_Input: HTMLInputElement;
    private cS_ContentSeparatorValue: string; 
    private separator: string;
    private editMode: boolean;
    private showLeft: boolean;
    private showLabel: boolean;
    private labeltext: string;
    private inputChange = (event: Event) => {
        let updatedvalue = (event.target as HTMLInputElement).value;
        let originalcontent = this.cS_ContentSeparatorValue.split(this.separator); 
        this.cS_ContentSeparatorValue = 
        this.showLeft ? 
        updatedvalue + " " + this.separator + " " + originalcontent[1].trim() : // Original Left (separator) New Value
        originalcontent[0].trim() + " " + this.separator + " " + updatedvalue;  // New Value (separator) Original Left
        this.notifyOutPutChanged();
    }
    //#endregion

    //#region Empty Constructor


    /**
     * Empty constructor.
     */
    constructor()
    {

    }

    //#endregion
    
    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
    {
        //#region - control initialization code 
        this.context = context;
        this.notifyOutPutChanged = notifyOutputChanged;
        this.state = state;
        this.container = container;
        //#endregion 
        this.loadData();
        this.loadForm();
    }
    /*
     * Used on load event to get the manifest data values. 
     */
    private loadData(): void {
        this.showLeft = this.context.parameters.LeftContent.raw;
        this.editMode = this.context.parameters.EditMode.raw; 
        this.separator = this.context.parameters.Separator.raw || ","; 
        this.cS_ContentSeparatorValue = this.context.parameters.ContentSeparatorValue.raw || "";   
        this.labeltext = this.context.parameters.LabelValue.raw || "";  
        this.showLabel = this.context.parameters.LabelDisplay.raw || false;        
    }
    /*
     * Used on load event to get the html control value and set the input html to the string. 
     */
    private loadForm(): void {
        this.createContainer();
        this.createLabel();
        this.createInput();
        this.setFormLoadValue();
    }
    /*
     * Used on load event to get the html control value and set the input html to the string. 
     */
    private createContainer(): void {
        this.cS_Container = this.getElement("div", "mycontainer", "mycontainer") as HTMLDivElement;
        this.container.appendChild(this.cS_Container);
    }
    /*
     * Used on load event to create the input control and append it to the Container. 
     */
    private createLabel(): void {
        try {
            this.cS_Label = this.getElement("label", "label", "mylabel") as HTMLLabelElement;
            let message = this.showLeft ? this.labeltext.split(this.separator)[0] : this.labeltext.split(this.separator)[1];
            this.cS_Label.innerText = "(" + message.trim() + ")";
            if(this.showLabel === true)
                this.container.appendChild(this.cS_Label);
        } catch (error){
            alert(error);
        }
    }
    /*
     * Used on load event to create the input control and append it to the Container. 
     */
    private createInput(): void {
        this.cS_Input = this.getElement("input", "Input", "myinput") as HTMLInputElement;
        this.cS_Input.disabled = !this.editMode;
        this.cS_Input.addEventListener("keyup", this.inputChange);
        this.container.appendChild(this.cS_Input);
    }
    private getElement(type: string, id: string, className: string): HTMLElement {
        let obj = document.createElement(type);
        obj.id = id;
        obj.className = className;
        return obj;
    }
    /**
     * Sets the Control value to the input for the Content Separator.     
     */
    private setFormLoadValue(): void {        
        try {
            let content = this.cS_ContentSeparatorValue.split(this.separator);   
            if (content.length < 2) return;     
            this.cS_Input.value = this.showLeft ? content[0].trim() : content[1].trim();
        } catch (error) {
            alert("Please contact support, the following error occurred: ERROR:" + error);
        }
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        // Add code to update control view
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs
    {
        return {
            ContentSeparatorValue : this.cS_ContentSeparatorValue
        };
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        // Add code to cleanup control if necessary
    }
}
