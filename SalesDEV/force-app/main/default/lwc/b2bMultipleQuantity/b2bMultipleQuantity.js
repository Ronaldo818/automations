import { LightningElement, track, api } from 'lwc';
import getMultipleQuantity from '@salesforce/apex/B2BMultipleQuantityController.getMultipleQuantity';

export default class B2bMultipleQuantity extends LightningElement {
    @api texto;
    @api recordId;
    @track pesoUnitario;

    connectedCallback() {
        console.log(this.recordId);
        getMultipleQuantity({productId: this.recordId})
        .then(result => {
            this.pesoUnitario = result;
        })
        .catch(error => {
            console.error('Erro no método "getMultipleQuantity"');
            console.error(error);
        });
    }
}