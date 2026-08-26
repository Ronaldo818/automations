import { api } from 'lwc';
import LightningModal from 'lightning/modal';

import { EFFECTIVE_ACCOUNT } from 'commerce/effectiveAccountApi'; 

import createMasterOrder from '@salesforce/apex/FvCreateMasterOrderController.createMasterOrder';
import getValidity from '@salesforce/apex/FvCreateMasterOrderController.getValidity';

import { refreshCartSummary } from 'commerce/cartApi';
import BASE_PATH from '@salesforce/community/basePath';

export default class FvCreateMasterOrderModal extends LightningModal {

    effectiveAccount = EFFECTIVE_ACCOUNT
    basePath = BASE_PATH

    isProcessing = true;
    created = false;
    cartId
    accountName = '';

    masterOrderId;
    masterOrderName;   

    cartItemsValidity = [];

    today;
    minDate;
    maxDate;
    endDate;
    
    _options
    @api
    set options(values){
        values.forEach(element => {
            this[element.name] = element.value
        });
        this._options = values;

        getValidity({cartId: this.cartId})
        .then((result) => {
            console.log('debug');
            this.cartItemsValidity = JSON.parse(result);
            this.today = this.cartItemsValidity.today;
            this.minDate = this.cartItemsValidity.minDate;
            this.maxDate = this.cartItemsValidity.maxDate;
            this.endDate = this.maxDate;
            this.isValid = this.cartItemsValidity.invalidItems.length === 0;
            this.isProcessing = false;
        })
        .catch((err) => {
            console.error('checkValidity: ', err);
            this.isProcessing = false;
        })
    }

    get effectiveAccountName(){
        return this.effectiveAccount?.accountName || this.accountName;
    }

    get options(){
        return this._options
    }

    get masterOrderUrl(){
        return this.basePath + '/masterorder/' + this.masterOrderId + '/detail';
    }

    handleComment(event){
        this.comment = event.detail.value;
    }

    handleCreateMasterOrder(){
        const allValid = [
            ...this.template.querySelectorAll('lightning-input, lightning-textarea'),
        ].reduce((validSoFar, inputCmp) => {
            inputCmp.setCustomValidity('');
            inputCmp.reportValidity();
            //console.log(inputCmp);
            this[inputCmp.name] = inputCmp.value;
            //console.log(inputCmp.name);   
            return validSoFar && inputCmp.checkValidity();
        }, true);
        if (allValid){
            this.isProcessing = true;
            createMasterOrder({ startDate: this.today, endDate: this.endDate, comment: this.comment, cart: this.cartDetails, items: this.cartItems})
            .then((result) => {
                refreshCartSummary();
                const r = JSON.parse(result);
                this.created = true;
                this.masterOrderId = r.id;
                this.masterOrderName = r.name;
                this.isProcessing = false;
            }).catch((err) => {
                console.error(err);
                this.isProcessing = false;
            });
        }
    }

    handleCloseModal(){
        this.close();
    }

}