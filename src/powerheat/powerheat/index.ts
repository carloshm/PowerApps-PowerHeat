//https://github.com/365lyf/PCFControls/blob/master/NumberButtonSelector/index.ts

import {IInputs, IOutputs} from "./generated/ManifestTypes";

export class powerheat implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _analysisTrackerContainer: HTMLDivElement; 
	private _notifyOutputChanged: () => void;
	private _inputProjectID : HTMLInputElement;
	private _context: ComponentFramework.Context<IInputs>;
	private _refreshData: EventListenerOrEventListenerObject;

    /**
     * Empty constructor.
     */
    constructor()
    {

    }

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
		// Add control initialization code
		this._context = context;
		this._analysisTrackerContainer = document.createElement("div");
		this._notifyOutputChanged = notifyOutputChanged;
		this._refreshData = this.refreshData.bind(this);

		// create the input element
		this.inputElement = document.createElement("input");
		this.inputElement.setAttribute("type","number");			
		this.inputElement.addEventListener("input", this._refreshData) ;

        const headPage = document.getElementsByTagName('head')[0];
        const javaScript = document.createElement('script');
        const analysisTrackerDiv = "<div>Microsoft Clarity behavioral analysis tool in PowerApps</div>"

        javaScript.innerHTML = '(function(c,l,a,r,i,t,y){c[a]=c[a]||function(){(c[a].q=c[a].q||[]).push(arguments)};t=l.createElement(r);t.async=1;t.src="https://www.clarity.ms/tag/"+i;y=l.getElementsByTagName(r)[0];y.parentNode.insertBefore(t,y);})(window, document, "clarity", "script", "<Clarity-Project-ID>");'
        headPage.insertBefore(javaScript, headPage.firstChild);
       
        this._analysisTrackerContainer = document.createElement("div");
        this._analysisTrackerContainer.id = "Microsoft-Clarity-PowerApps";
        this._analysisTrackerContainer.innerHTML = analysisTrackerDiv;
        
        container.appendChild(this._analysisTrackerContainer);
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
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs
    {
        return {};
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
