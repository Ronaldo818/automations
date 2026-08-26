import { api, LightningElement } from 'lwc';
import FORM_FACTOR from '@salesforce/client/formFactor';
import getMasterOrders from '@salesforce/apex/MasterOrderPickerController.getMasterOrders';
import { effectiveAccount } from 'commerce/effectiveAccountApi'; 
import getExistingCartItem from '@salesforce/apex/MasterOrderPickerController.getExistingCartItem';

const DEFAULT_OPTIONS = [
    {
        value: 'null', label: 'Não se aplica', labelStruct: {
            name: 'Não se aplica',
            aQty: null,
            nPrice: null
        }
    }]

export default class FvMasterOrderPicker extends LightningElement {
    
    options = DEFAULT_OPTIONS;
    value = 'null'
    selectedQuantity = 0;
    init = false;
    disabled = true;

    _effectiveAccountId;

    @api
    set effectiveAccountId(value){
        if(value){
            this._effectiveAccountId = value;
            this.getData();
        }
    }
    @api
    cartId;

    get effectiveAccountId(){
        return this._effectiveAccountId || effectiveAccount.accountId;
    }

    _productId
    @api
    set productId(value) {
        //console.log('set productId:', value);
        if (value && this._productId !== value) {
            //console.log('enter');
            this._productId = value;
            this.getData();
        }
        else {
            //console.log('not enter');
            //console.log(value, this._productId);
        }
    }

    @api
    update(){
        this.init = true;
        this.getData();
    }

    get isReady(){
        return this._productId && this.effectiveAccountId && !this.isGettingData;
    }

    isGettingData = false;

    async getData() {
        if(this.isReady){
            //console.log(this._effectiveAccountId, this._productId, this.cartId);
            this.isGettingData = true;
            Promise.all([getMasterOrders({ effectiveAccountId: this.effectiveAccountId, productId: this._productId }), getExistingCartItem({ productId: this._productId, cartId: this.cartId, accId: this.effectiveAccountId })]).then((values) => {
                this.options = DEFAULT_OPTIONS;
                const tmp = JSON.parse(values[0]);
                //console.log(tmp);
                this.options = this.options.concat(tmp.map((item) => ({
                    label: item.MasterOrderId__r.Name + ' | R$' + item.NegotiatedUnitPriceKg__c.toFixed(2).replace('.',',') + ' | Disp: ' + item.AvailableQuantity__c,
                    labelStruct: {
                        name: item.MasterOrderId__r.Name,
                        aQty: item.AvailableQuantity__c,
                        nPrice: item.NegotiatedUnitPriceKg__c
                    },
                    value: item.Id
                })));
                if (this.isMobile) {
                    let elem = this.template.querySelector('[data-id="masterOrderSelect"]');
                    //console.log(elem);
                    const tempOptions = document.createDocumentFragment();
                    this.options.forEach(option => {
                        const optValue = option.value;
                        const optLabel = option.label;

                        const opt = document.createElement('option');
                        opt.value = optValue;
                        opt.textContent = optLabel;

                        tempOptions.appendChild(opt);
                    });
                    elem.appendChild(tempOptions);
                }
                if(values[1]){
                    const tmp = JSON.parse(values[1]);
                    this.value = tmp.MasterOrderItemId__c ? tmp.MasterOrderItemId__c : 'null';
                    this.selectedQuantity = tmp.Quantity;
                    this.nPrice = tmp.UnitAdjustedPriceWithItemAdj;
                    if(this.value != 'null' && this.isMobile){
                        this.template.querySelector('[data-id="masterOrderSelect"] [value="'+ this.value +'"]').setAttribute("selected", "");
                    }
                    this.init = true
                }
                else{
                    this.init = true
                    this.selectedQuantity = 0;
                    this.nPrice = null;
                }
                this.handleDispatchEvent();
                this.isGettingData = false;
                this.disabled = this.options.length === 1;
            }).catch((err) => {
                this.init = true
                this.handleDispatchEvent();
                this.isGettingData = false;
                this.disabled = true;
                console.error(err);
            });
        }
    }

    get productId() {
        return this._productId
    }

    get isMobile() {
        return FORM_FACTOR === 'Small';
    }

    async handleSelect() {
        if (this.isMobile) {
            //this.value = event.target.selectedOptions[0].value;
            this.value = event.target.value;
        }
        else {
            this.value = event.detail.value;
        }
        this.handleDispatchEvent();
    }

    handleDispatchEvent(){
        let selectedOption;
        const optionsLength = this.options.length;
        for (let i = 0; i < optionsLength; i++) {
            if (this.options[i].value === this.value) {
                selectedOption = this.options[i];
                break;
            }
        }
        // Creates the event with the contact ID data.
        const selectedEvent = new CustomEvent("selected", { detail: { value: this.value, maxQty: selectedOption?.labelStruct?.aQty, nPrice: selectedOption?.labelStruct?.nPrice?.toFixed(2) || this.nPrice?.toFixed(2), qty: this.selectedQuantity, init: this.init } });
        this.init = false
        //event.preventDefault();

        // Dispatches the event.
        this.dispatchEvent(selectedEvent);
    }

}