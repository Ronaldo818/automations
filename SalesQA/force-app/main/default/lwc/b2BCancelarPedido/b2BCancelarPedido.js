import { LightningElement, track, api } from 'lwc';

export default class B2BCancelarPedido extends LightningElement {
    @api recordId;
    @api alertMesage;
    @api comercialNumber;
    @track showAlert = false;
    @track isLoading = false;

    cancelOrder() {
        this.showAlert = true;
    }

    closeAlert() {
        this.showAlert = false;    
    }
}