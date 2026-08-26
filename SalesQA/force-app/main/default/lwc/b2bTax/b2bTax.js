import { LightningElement, api, track, wire } from 'lwc';
import checkoutMessage from '@salesforce/messageChannel/checkoutMessage__c';
import {publish, MessageContext} from "lightning/messageService";
import cartChanged from "@salesforce/messageChannel/lightning__commerce_cartChanged";
import calculaImposto from '@salesforce/apex/B2BCheckoutController.calculaImposto';
import {getRecord} from 'lightning/uiRecordApi';
import PROFILE_NAME_FIELD from '@salesforce/schema/User.Profile.Name';
import strUserId from '@salesforce/user/Id';

export default class B2bTax extends LightningElement {

    @wire(MessageContext)
    messageContext;

    @api title;
    @api minimumOrderValue;
    @api messageMinimumOrderError;
    @track isLoading;
    @track __subtotal = 0;
    @track __promocao = 0;
    @track __impostos = 0;
    @track __valorTotal = 0;
    
    @track showAlert = false;
    @track showAlertMesage;
    @track alertLogId;
    @track prfName;

    @wire(getRecord, {recordId: strUserId, fields: [PROFILE_NAME_FIELD]}) 
    wireuser({error, data}) {
        if(error) {
            console.log('Get User Profile ERROR => ' + error);
        }
        else if(data) {
            this.prfName = data.fields.Profile.value.fields.Name.value;
            console.log('prfName => ' + this.prfName);  
        }
    }

    connectedCallback() {
        this.isLoading = true;
        calculaImposto().then(result => {
            if(result) {
                let retorno = JSON.parse(result);
                if(retorno.hasError) {
                    this.showAlert = true;
                    this.showAlertMesage = retorno.errorMessage;
                    this.alertLogId = retorno.logId;
                    this.isLoading = false;
                }
                else {
                    let cartSummary = retorno.cartSum;
                    if(cartSummary.grandTotalAmount > parseInt(this.minimumOrderValue)) {
                        this.showAlert = false;
                        const payload = { taxSuccess: true };
                        publish(this.messageContext, checkoutMessage, payload);
                        this.showAlertMesage = '';
                        this.__subtotal = cartSummary.totalProductAmount;
                        this.__promocao = cartSummary.totalPromotionalAdjustmentAmount;
                        this.__impostos = cartSummary.totalTaxAmount;
                        this.__valorTotal = cartSummary.grandTotalAmount;
                        this.isLoading = false;
                    }
                    else if(cartSummary.grandTotalAmount > 0 && cartSummary.grandTotalAmount < parseInt(this.minimumOrderValue)) {
                        this.__subtotal = cartSummary.totalProductAmount;
                        this.__promocao = cartSummary.totalPromotionalAdjustmentAmount;
                        this.__impostos = cartSummary.totalTaxAmount;
                        this.__valorTotal = cartSummary.grandTotalAmount;
                        this.showAlert = true;
                        this.showAlertMesage = this.messageMinimumOrderError;
                        this.isLoading = false;
                    }
                    else {
                        this.showAlert = true;
                        this.showAlertMesage = 'Ocorreu um erro inesperado no calculo de impostos, favor tentar novamente mais tarde ou entrar em contato com o suporte.';
                        this.isLoading = false;
                    }
                }
            }

            setTimeout(function(){
                this.dispatchEvent(new CustomEvent("cartchanged", {
                    bubbles: true,
                    composed: true
                }));
                publish(this.messageContext, cartChanged);
            }, 1000);
            setTimeout(function(){
                this.dispatchEvent(new CustomEvent("cartchanged", {
                    bubbles: true,
                    composed: true
                }));
                publish(this.messageContext, cartChanged);
            }, 2000);
            setTimeout(function(){
                this.dispatchEvent(new CustomEvent("cartchanged", {
                    bubbles: true,
                    composed: true
                }));
                publish(this.messageContext, cartChanged);
            }, 5000);
            setTimeout(function(){
                this.dispatchEvent(new CustomEvent("cartchanged", {
                    bubbles: true,
                    composed: true
                }));
                publish(this.messageContext, cartChanged);
            }, 10000);
            setTimeout(function(){
                this.dispatchEvent(new CustomEvent("cartchanged", {
                    bubbles: true,
                    composed: true
                }));
                publish(this.messageContext, cartChanged);
            }, 15000);
            setTimeout(function(){
                this.dispatchEvent(new CustomEvent("cartchanged", {
                    bubbles: true,
                    composed: true
                }));
                publish(this.messageContext, cartChanged);
            }, 20000);
        })
        .catch(error => {
            this.showAlert = true;
            this.showAlertMesage = JSON.stringify(error);
            this.isLoading = false;
        })
    }

    closeAlert() {
        this.isLoading = true;
        this.showAlertMesage = '';
        this.showAlert = false;
        this.isLoading = false;
        if(this.prfName.toLowerCase() !== 'system administrator' && this.prfName.toLowerCase() !== 'administrador do sistema') {
            window.location.href = window.location.origin + '/cart';
        }
    }
}