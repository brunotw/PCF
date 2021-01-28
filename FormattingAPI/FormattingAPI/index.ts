import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class FormattingAPI implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	// PCF framework delegate that will be assigned to this object that would be called whenever any update happens. 
	private _notifyOutputChanged: () => void;
	// Reference to the div element that holds together all the HTML elements that you are creating as part of this control
	private divElement: HTMLDivElement;
	// Reference to HTMLTableElement rendered by control
	private _tableElement: HTMLTableElement;
	// Reference to the control container HTMLDivElement
	// This element contains all elements of your custom control example
	private _container: HTMLDivElement;
	// Reference to ComponentFramework Context object
	private _context: ComponentFramework.Context<IInputs>;
	// Flag if control view has been rendered
	private _controlViewRendered: Boolean;

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Dataset values are not initialized here, use updateView.
	 * @param context The entire property bag is available to control through the Context Object; It contains values as set up by the customizer and mapped to property names that are defined in the manifest and to utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a control's life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='starndard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
		this._notifyOutputChanged = notifyOutputChanged;
		this._controlViewRendered = false;
		this._context = context;

		this._container = document.createElement("div");
		this._container.classList.add("TSFormatting_Container");
		container.appendChild(this._container);
	}

	public updateView(context: ComponentFramework.Context<IInputs>): void {
		if (!this._controlViewRendered) {
			// Render and add HTMLTable to the custom control container element
			let tableElement: HTMLTableElement = this.createHTMLTableElement();
			this._container.appendChild(tableElement);

			this._controlViewRendered = true;
		}
	}


	public getOutputs(): IOutputs { return {}; }

	public destroy() { }

	/** 
* Creates an HTML Table that showcases examples of basic methods that are available to the custom control
* The left column of the table shows the method name or property that is being used
* The right column of the table shows the result of that method name or property
*/
	private createHTMLTableElement(): HTMLTableElement {
		// Create HTML Table Element
		let tableElement: HTMLTableElement = document.createElement("table");
		tableElement.setAttribute("class", "FormattingControlSampleHtmlTable_HtmlTable");

		// Create header row for table
		let key: string = "Example Method";
		let value: string = "Result";
		tableElement.appendChild(this.createHTMLTableRowElement(key, value, true));

		// Example use of formatCurrency() method 
		// Change the default currency and the precision or pass in the precision and currency as additional parameters.
		key = "formatCurrency()";
		value = this._context.formatting.formatCurrency(10250030);
		tableElement.appendChild(this.createHTMLTableRowElement(key, value, false));

		// Example use of formatDecimal() method 
		// Change the settings from user settings to see the output change its format accordingly
		key = "formatDecimal()";
		value = this._context.formatting.formatDecimal(123456.2782);
		tableElement.appendChild(this.createHTMLTableRowElement(key, value, false));

		// Example use of formatInteger() method
		// change the settings from user settings to see the output change its format accordingly.
		key = "formatInteger()";
		value = this._context.formatting.formatInteger(12345);
		tableElement.appendChild(this.createHTMLTableRowElement(key, value, false));

		// Example use of formatLanguage() method
		// Install additional languages and pass in the corresponding language code to see its string value
		key = "formatLanguage()";
		value = this._context.formatting.formatLanguage(1033);
		tableElement.appendChild(this.createHTMLTableRowElement(key, value, false));

		// Example of formatDateYearMonth() method
		// Pass a JavaScript Data object set to the current time into formatDateYearMonth method to format the data
		// and get the return in Year, Month format
		key = "formatDateYearMonth()";
		value = this._context.formatting.formatDateYearMonth(new Date());
		tableElement.appendChild(this.createHTMLTableRowElement(key, value, false));

		// Example of getWeekOfYear() method
		// Pass a JavaScript Data object set to the current time into getWeekOfYear method to get the value for week of the year
		key = "getWeekOfYear()";
		value = this._context.formatting.getWeekOfYear(new Date()).toString();
		tableElement.appendChild(this.createHTMLTableRowElement(key, value, false));

		return tableElement;
	}

	/**
	 * Helper method to create an HTML Table Row Element
	 * @param key : string value to show in left column cell
	 * @param value : string value to show in right column cell
	 * @param isHeaderRow : true if method should generate a header row
	 */
	private createHTMLTableRowElement(key: string, value: string, isHeaderRow: boolean): HTMLTableRowElement {
		let keyCell: HTMLTableCellElement = this.createHTMLTableCellElement(key, "FormattingControlSampleHtmlTable_HtmlCell_Key", isHeaderRow);
		let valueCell: HTMLTableCellElement = this.createHTMLTableCellElement(value, "FormattingControlSampleHtmlTable_HtmlCell_Value", isHeaderRow);

		let rowElement: HTMLTableRowElement = document.createElement("tr");
		rowElement.setAttribute("class", "FormattingControlSampleHtmlTable_HtmlRow");
		rowElement.appendChild(keyCell);
		rowElement.appendChild(valueCell);

		return rowElement;
	}

	/**
	 * Helper method to create an HTML Table Cell Element
	 * @param cellValue : string value to inject in the cell
	 * @param className : class name for the cell
	 * @param isHeaderRow : true if method should generate a header row cell
	 */
	private createHTMLTableCellElement(cellValue: string, className: string, isHeaderRow: boolean): HTMLTableCellElement {
		let cellElement: HTMLTableCellElement;

		if (isHeaderRow) {
			cellElement = document.createElement("th");
			cellElement.setAttribute("class", "FormattingControlSampleHtmlTable_HtmlHeaderCell " + className);
			let textElement: Text = document.createTextNode(cellValue);
			cellElement.appendChild(textElement);
		}
		else {
			cellElement = document.createElement("td");
			cellElement.setAttribute("class", "FormattingControlSampleHtmlTable_HtmlCell " + className);
			let textElement: Text = document.createTextNode(cellValue);
			cellElement.appendChild(textElement);
		}
		return cellElement;
	}


}