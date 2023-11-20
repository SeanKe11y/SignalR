import {IInputs, IOutputs} from "./generated/ManifestTypes";
import * as signalR from "@microsoft/signalr";
import { IMessage } from "./IMessage";

export class NotificationComponent implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;
    private _connection?: signalR.HubConnection;
    private _signalRApi: string;
    private _message: IMessage;

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
        this._notifyOutputChanged = notifyOutputChanged;
        this._signalRApi = context.parameters.signalRHubConnectionUrl.raw
            ? context.parameters.signalRHubConnectionUrl.raw
            : "";
        
        this._connection = new signalR.HubConnectionBuilder()
            .withUrl(this._signalRApi)
            .configureLogging(signalR.LogLevel.Information)
            .withAutomaticReconnect()
            .build();
        
        //configure the event when a new message arrives
        this._connection.on("newMessage", (message:IMessage) => {
            this._message = message;
            this._notifyOutputChanged();
        });

        //connect
        this._connection
            .start()
            .catch(err => {
                console.log(err);
                this._connection!.stop();
            });
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        //When the messageData is updated this code will run and we send the message to signalR
        this._context = context;
        //check for updateProperties 	
		if (this._context.updatedProperties != null && this._context.updatedProperties.length != 0) {
            if (this._context.updatedProperties.indexOf("messageData") > -1) {
                let messageData = JSON.parse(this._context.parameters.messageData.raw!= null?
                    this._context.parameters.messageData.raw:"");

                if(messageData.text) {
                    this.httpCall(messageData, (res)=>{ console.log(res)});
                }
            }
        }        
    }

    public httpCall(data: IMessage, callback:(result:any)=>any): void {
        var xhr = new XMLHttpRequest();
        xhr.open("post", this._signalRApi + "/messages", true);
        if (data != null) {
            xhr.setRequestHeader('Content-Type', 'application/json');
            xhr.send(JSON.stringify(data));
        }
        else xhr.send();
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs
    {
        //This code will run when we call notifyOutputChanged when we receive a new message
        //here is where the message gets exposed to outside
        let output: IOutputs = {
            messageReceivedText: this._message.text,
            messageReceivedTime: this._message.time,
            messageReceivedSender: this._message.sender,
            messageReceivedImage: this._message.image
        };
        return output;
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        // Add code to cleanup control if necessary
        if (this._connection) {
            this._connection.stop();
        }
    }
}
