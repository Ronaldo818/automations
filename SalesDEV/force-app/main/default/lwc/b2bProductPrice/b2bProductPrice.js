import { LightningElement, api, track, wire } from 'lwc';
import getPrice from '@salesforce/apex/B2BProductPriceController.getProductPrice';

export default class B2bProductPrice extends LightningElement {
    @api recordId;
    @api fieldLabel;
    @track price;
    @track error;
    @track isLoading = true;

    connectedCallback() {
        getPrice({productId: this.recordId})
        .then(result => {
            this.price = result;
            this.isLoading = false;
        })
        .catch(error => {
            this.error = JSON.stringify(error);
            console.log('========== error');
            console.error(this.error);
            this.isLoading = false;
        });
    }
}