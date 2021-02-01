import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class PCFTextComponent implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	//Value of the field is stored and used inside the control
	private _value: string;

	//Reference to ComponentsFramework Context object
	private _context: ComponentFramework.Context<IInputs>;

	//Reference to the control container HTMLDivElement
	//This element contains all elements of our custom control example
	private _container: HTMLDivElement;

	constructor() {

	}

	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
		// Add control initialization code
		this._context  = context;
		this._container = document.createElement("div");
		this._value = context.parameters.sampleProperty.raw ? context.parameters.sampleProperty.raw : "";

		this._container.innerText = this._value;
		container.appendChild(this._container);
	}

	public updateView(context: ComponentFramework.Context<IInputs>): void {
		// Add code to update control view
		this._value = context.parameters.sampleProperty.raw ? context.parameters.sampleProperty.raw : "";
		this._context = context;
		this._container.innerText = this._value;	
	}


	public getOutputs(): IOutputs {
		return {};
	}

	public destroy(): void {
		// Add code to cleanup control if necessary
	}
}