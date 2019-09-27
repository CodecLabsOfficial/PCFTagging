import { IInputs, IOutputs } from "./generated/ManifestTypes";
// Import & attach jQuery to Window
import * as $ from "jquery";
declare var window: any;
window.$ = window.jQuery = $;

// @ts-ignore
import * as autocomplete from "jquery-ui/ui/widgets/autocomplete";
import "jquery-ui";
import "bootstrap";
import "bootstrap-tagsinput";
import * as DynamicsWebApi from "dynamics-web-api";

interface JQuery {
    autocomplete(config: { source: string[]; }): any;
}

export class CCFTagging implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    /** HTML elements */
    private _tagsInput: HTMLElement;

    /** Properties */
    private _tagsString: string;
    private _recordId: string;
    private _recordLogicalName: string;
    private _recordName?: string;
    private _webApi: DynamicsWebApi;
    private _typeAheadSource: Array<string>;

    /** Events */
    private _tagAdded: EventListenerOrEventListenerObject;
    private _tagRemoved: EventListenerOrEventListenerObject;

    /** General */
    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;
    private _container: HTMLDivElement;

	/**
	 * Empty constructor.
	 */
    constructor() {

    }

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='starndard', it will receive an empty div element within which it can render its content.
	 */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
        // Assigning environment variables.
        this.initVars(context, notifyOutputChanged, container);
        // Register eventhandler functions 
        this.initEventHandlers();
        // Tags input
        this.initTagInput();
        // add to the container so that it renders on the UI. 
        this._container.appendChild(this._tagsInput);

        this.getRecordInfo();
        this.initTagsValue();

        var initAC = autocomplete;
    }

    private initVars(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, container: HTMLDivElement): void {
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._container = container;
        this._webApi = new DynamicsWebApi({ webApiVersion: '9.0' });
        this._tagsString = "";
    }

    private initEventHandlers(): void {

    }

    private initTagInput(): void {
        this._tagsInput = document.createElement("input");
        this._tagsInput.setAttribute("type", "text");
        this._tagsInput.setAttribute("id", "txtTags");
        this._tagsInput.setAttribute("class", "form-control");
        this._tagsInput.dataset.role = "tagsinput";
    }

    public initTagsValue(): void {
        this.loadTags(() => {
            this._tagsInput.setAttribute("value", this._tagsString.toString());

            this.loadTypeAhead();
        });
    }

    private loadTypeAhead(): void {
        this._webApi.retrieveMultiple("clabs_tags", ["clabs_tag"]).then((records) => {

            let params = {
                'onAddTag': (input: any, value: string) => { this.tagAdded(value); },
                'onRemoveTag': (input: any, value: string) => { this.tagRemoved(value); }
            }

            if (records.value && records.value.length > 0) {
                this._typeAheadSource = $.map(records.value, (item) => { return item.clabs_tag; });

                // @ts-ignore
                params.autocomplete = { source: this._typeAheadSource };
            }

            // @ts-ignore
            $(this._tagsInput).tagsInput(params);
        }).catch((error) => {
            console.error(error);
            debugger;
        });
    }

    private loadTags(callback: () => void): void {
        let filter = "clabs_id eq '" + this._recordId + "'";

        let request = this.generateRequest("clabs_taggedrecords", ["clabs_taggedrecordid,_clabs_tagid_value"], filter, true);
        // @ts-ignore
        request.expand = "clabs_tagid($select=clabs_tag)";

        let daddy = this;
        this._webApi.retrieveMultipleRequest(request).then(function (records) {
            if (records.value && records.value.length > 0) {
                daddy.formatTags(records.value);
            }

            callback();
        }).catch((error: any) => {
            console.error(error);
            debugger;
        });


    }

    private formatTags(value: any): void {
        if (value && value.length > 0) {
            $.each(value, (index, item) => {
                this._tagsString += item.clabs_tagid.clabs_tag + ",";
            });
        }
    }

    public getRecordInfo(): void {
        this._recordId = Xrm.Page.data.entity.getId().replace("{", "").replace("}", "").toLowerCase();
        this._recordLogicalName = Xrm.Page.data.entity.getEntityReference().entityType;
        this._recordName = Xrm.Page.data.entity.getEntityReference().name;
    }

    public tagAdded(addedTag: string): void {
        this.updateTagField();

        // @ts-ignore
        let changedTag = addedTag;

        var daddy = this;
        this.retrieveTagId(changedTag, (tagId: string | null) => {
            if (tagId == null) {
                daddy.createTag(changedTag, (createdTagId: string) => {
                    daddy.createTaggedRecord(createdTagId);
                });
            } else {
                daddy.createTaggedRecord(tagId);
            }
        });
    }

    private createTag(changedTag: string, callback: (recordId: string) => void): void {
        var tag = {
            clabs_tag: changedTag.toLowerCase()
        };

        this._webApi.create(tag, "clabs_tags").then(function (id: string) {
            callback(id);
        }).catch((error) => {
            console.error(error);
            debugger;
        });
    }

    private retrieveTagId(changedTag: string, callback: (recordId: string | null) => void): void {
        let filter = "clabs_tag eq '" + changedTag.toLowerCase() + "'";

        let request = this.generateRequest("clabs_tags", ["clabs_tagid"], filter, true);

        this._webApi.retrieveMultipleRequest(request).then(function (records) {
            if (records.value && records.value.length > 0) {
                callback(records.value[0].clabs_tagid);
            } else {
                callback(null);
            }
        }).catch((error: any) => {
            console.error(error);
            debugger;
        });
    }

    private createTaggedRecord(tagId: string): void {
        var taggedRecord = {
            clabs_logicalname: this._recordLogicalName,
            clabs_name: this._recordName,
            clabs_id: this._recordId
        };

        let daddy = this;
        this._webApi.create(taggedRecord, "clabs_taggedrecords").then((id: string) => {
            var createdTagged = {};

            // @ts-ignore
            createdTagged["_clabs_tagid_value@odata.bind"] = "/clabs_tags(" + tagId + ")";

            daddy._webApi.associateSingleValued("clabs_taggedrecords", id, "clabs_tagid", "clabs_tags", tagId);
        }).catch((error: any) => {
            console.error(error);
            debugger;
        });
    }

    public tagRemoved(removedTag: string): void {
        this.updateTagField();

        // @ts-ignore
        let changedTag = removedTag;

        var daddy = this;
        this.retrieveTagId(changedTag, function (tagId: string | null) {
            daddy.retrieveTaggedRecordId(tagId as string, function (recordId: string) {
                daddy.deletedTaggedRecord(recordId);
            });
        });
    }

    private retrieveTaggedRecordId(tagId: string, callback: (recordId: string) => void): void {
        let filter = "_clabs_tagid_value eq " + tagId.toLowerCase() + " and clabs_id eq '" + this._recordId + "'";

        let request = this.generateRequest("clabs_taggedrecords", ["clabs_taggedrecordid"], filter, true);

        this._webApi.retrieveMultipleRequest(request).then(function (records) {
            if (records.value && records.value.length > 0) {
                callback(records.value[0].clabs_taggedrecordid);
            }
        }).catch((error: any) => {
            console.error(error);
            debugger;
        });
    }

    public generateRequest(logicalName: string, select: string[], filter: string, async: boolean): object {
        let request = {
            collection: logicalName,
            select: select,
            filter: filter,
            async: async
        };

        return request;
    }

    private deletedTaggedRecord(taggedRecordId: string): void {
        this._webApi.deleteRecord(taggedRecordId, "clabs_taggedrecords").then(function () {

        }).catch((error: any) => {
            console.error(error);
            debugger;
        });
    }

    public updateTagField(): void {
        // ts-ignore
        //var crmTagsInput = this._context.parameters.Tags.attributes.LogicalName;
        //var tagsValue = $(this._tagsInput).val();
        //Xrm.Page.getAttribute(crmTagsInput).setValue(tagsValue);
    }

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
    public updateView(context: ComponentFramework.Context<IInputs>): void {

    }

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
    public getOutputs(): IOutputs {
        return {
            Tags: this._tagsString
        };
    }

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }
}