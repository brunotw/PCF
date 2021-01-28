import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { FacepileBasicExample, IFacepileBasicExampleProps } from "./Facepile";

export class ReactStandardControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	// reference to the notifyOutputChanged method
	private notifyOutputChanged: () => void;
	// reference to the container div
	private theContainer: HTMLDivElement;
	// reference to the React props, prepopulated with a bound event handler
	private props: IFacepileBasicExampleProps = {
		numberFacesChanged: this.numberFacesChanged.bind(this)
	};

	constructor() { }

	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
		this.notifyOutputChanged = notifyOutputChanged;
		this.props.numberOfFaces = context.parameters.numberOfFaces.raw || 3;
		this.theContainer = container;
	}

	public updateView(context: ComponentFramework.Context<IInputs>): void {
		if (context.updatedProperties.includes("numberOfFaces"))
			this.props.numberOfFaces = context.parameters.numberOfFaces.raw || 3;

		// Render the React component into the div container
		ReactDOM.render(
			// Create the React component
			React.createElement(FacepileBasicExample, // the class type of the React component found in Facepile.tsx
				this.props
			),
			this.theContainer
		);
	}

	/**
	 * Called by the React component when it detects a change in the number of faces shown
	 * @param newValue The newly detected number of faces
	 */
	private numberFacesChanged(newValue: number) {
		// only update if the number of faces has truly changed
		if (this.props.numberOfFaces !== newValue) {
			this.props.numberOfFaces = newValue;
			this.notifyOutputChanged();
		}
	}

	/**
	 * It is called by the framework prior to a control receiving new data.
	 * @returns an object based on nomenclature that is defined in the manifest, expecting object[s] for property marked as "bound" or "output"
	 */
	public getOutputs(): IOutputs {
		return {
			numberOfFaces: this.props.numberOfFaces
		};
	}

	/**
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup,
	 * for example, canceling any pending remote calls, removing listeners, and so on.
	 */
	public destroy(): void {
		ReactDOM.unmountComponentAtNode(this.theContainer);
	}
}