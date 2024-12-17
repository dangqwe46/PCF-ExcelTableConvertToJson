import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as XLSX from "xlsx";
export class ExcelUploader
  implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
  private _container: HTMLDivElement;
  private _fileInput: HTMLInputElement;
  private _notifyOutputChanged: () => void;
  private _outputType: string;
  private _value: string;
  private _fileName: string;
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
  ): void {
    this._notifyOutputChanged = notifyOutputChanged;

    this._value = "";
    this._fileName = "";
    this._fileInput = document.createElement("input");
    this._fileInput.type = "file";
    this._fileInput.id = "fileInput";
    this._fileInput.accept = ".xls,.xlsx";
    this._fileInput.style.display = "none"; // Hide the input

    // Create a label that will act as the button
    const label = document.createElement("label");
    label.htmlFor = "fileInput"; // Link the label to the input
    label.className = "custom-file-upload"; // Add a custom class for styling
    label.innerHTML = `<svg width="20" height="20" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <path d="M13.0438 5.52673L5.29382 13.407C5.03724 13.6901 4.89939 14.0611 4.90879 14.4431C4.9182 14.8251 5.07414 15.1889 5.34434 15.4591C5.61454 15.7293 5.9783 15.8852 6.3603 15.8946C6.7423 15.904 7.11329 15.7662 7.39646 15.5096L16.635 6.14078C17.1482 5.57444 17.4239 4.83246 17.4051 4.06846C17.3863 3.30446 17.0744 2.57695 16.534 2.03655C15.9936 1.49615 15.2661 1.18426 14.5021 1.16545C13.7381 1.14664 12.9961 1.42235 12.4298 1.9355L3.19118 11.3043C2.3547 12.1408 1.88477 13.2753 1.88477 14.4583C1.88477 15.6413 2.3547 16.7758 3.19118 17.6123C4.02766 18.4487 5.16217 18.9187 6.34514 18.9187C7.5281 18.9187 8.66261 18.4487 9.4991 17.6123L17.1374 9.99251" stroke="white" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
                        </svg>  Choose File`; // Button text
    // Bind the onchange event
    this._fileInput.onchange = this.handleFile.bind(this);
    // Append the input and label to the container
    container.appendChild(this._fileInput);
    container.appendChild(label);

    // this._notifyOutputChanged();
  }

  /**
   * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
   */
  public updateView(context: ComponentFramework.Context<IInputs>): void {
    // Add code to update control view
  }

  /**
   * It is called by the framework prior to a control receiving new data.
   * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
   */
  public getOutputs(): IOutputs {
    return {
      jsonOutput: JSON.stringify(this._value), // Example output
      fileName: this._fileName,
    };
  }

  /**
   * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
   * i.e. cancelling any pending remote calls, removing listeners, etc.
   */
  public destroy(): void {
    // Add code to cleanup control if necessary
  }
  private handleFile(event: Event): void {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (file) {
      const reader = new FileReader();
      this._fileName = file.name;
      reader.onload = (e) => {
        const data = new Uint8Array(e.target!.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });
        const jsonOutput = this.convertExcelToJson(workbook);
        this._value = jsonOutput;
        // console.log(jsonOutput); // Handle the JSON output as required
        // Optionally notify output change
        this._fileInput.value = "";
        this._notifyOutputChanged();
      };
      reader.readAsArrayBuffer(file);
    }
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private convertExcelToJson(workbook: XLSX.WorkBook): string {
    const jsonOutput: { [sheetName: string]: string[] } = {};
    workbook.SheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      jsonOutput[sheetName] = XLSX.utils.sheet_to_json(worksheet);
    });
    return JSON.stringify(jsonOutput);
  }
}
