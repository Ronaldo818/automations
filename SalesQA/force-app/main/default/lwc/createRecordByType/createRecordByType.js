import { api, LightningElement, wire } from 'lwc';
import getRecordTypes from '@salesforce/apex/RecordTypePickerController.getRecordTypes';


import recordTypePicker from './createRecordByType.html';
import errorTemplate from './createRecordByTypeError.html';

export default class RecordTypePicker extends LightningElement {

    @api get selectedRecordTypeId(){
        return this._selectedRecordType;
    }
    _selectedRecordType;

    @api get availableRecordTypes(){
        return this._returnedRecordTypes;
    }
    _returnedRecordTypes;

    @api objectApiName;

    @api
    get label(){
        return this._label;
    }
    set label(val){
        this._label = val;
    }
    _label = 'Select a Record Type';

    @api
    get sortOrder(){
        return this._sortOrder;
    }
    set sortOrder(val){
        this._sortOrder = val;
    }
    _sortOrder = 'ASC';

    @api
    get mode(){
        return this._mode;
    }
    set mode(val){
        this._mode = val?.toLowerCase();
    }
    _mode = 'live';

    _error;
    _selectedValue;

    get isPreview() {
        return this._mode === 'preview';
    }

    render(){
        if (this._error) {
            return errorTemplate;
        }else{
            return recordTypePicker;
        }
    }

    connectedCallback(){
        if (this.recordTypes.data && this.recordTypes.data.length === 0){
            this._error = 'You must select an Object that has active Record Types';
        }
        const path = window.location.split('/');
        for(let i = 0; i < path.length ;i++){
            if(path[i] == 'o'){
                this.sObjectApiName = path[i+1];
                break;
            }
        }
    }

    @wire(getRecordTypes, { sObjectApiName : '$objectApiName'})
    recordTypes({ error,data }) {
        if (data) {
            // we'll use this to invert the order if _orderBy is set to descending, default to ascending
            const ORDER_INVERTER = this._orderBy === 'DESC' ? -1 : 1; 
            let returnData = [...data];
            this._returnedRecordTypes = returnData.sort(function(a,b) {
                const aName = a.Name.toLowerCase(); //Object.assign({},a).Name;
                const bName = b.Name.toLowerCase(); //Object.assign({},b).Name;

                if (aName < bName) {
                    return -1 * ORDER_INVERTER;
                }

                if (aName > bName) {
                    return 1 * ORDER_INVERTER;
                }

                return 0;
            });
            this._error = undefined;
        }
        if (error) {
            this._error = error;
        }
    }

    
    onChange(event){
        this._selectedRecordType = event.target.value;
        console.log('this._selectedRecordType :'+this._selectedRecordType );
    }
}