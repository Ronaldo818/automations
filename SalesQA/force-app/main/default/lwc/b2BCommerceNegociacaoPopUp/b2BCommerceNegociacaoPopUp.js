import { LightningElement, api, track, wire } from 'lwc';
import getNegociacao from '@salesforce/apex/B2BAgrosysNegociacao.getCustomerCode'
export default class B2BCommerceNegociacaoPopUp extends LightningElement {
    @api message;
    @track showAlert = false;

    connectedCallback() {
        getNegociacao()
        .then( result => {
            this.showAlert = result;
        })
        .catch( error => {
            console.log('ERROR connectedCallback ==>');
            console.log(error);
            this.showAlert = false;
        });
    }

    closeAlert() {
        this.showAlert = false;
    }
}