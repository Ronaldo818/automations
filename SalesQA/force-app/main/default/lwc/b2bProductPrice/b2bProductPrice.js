import { LightningElement, api, track, wire } from 'lwc';
import getPrice from '@salesforce/apex/B2BProductPriceController.getProductPrice';
import { getRecord } from 'lightning/uiRecordApi';

export default class B2bProductPrice extends LightningElement {
    @api recordId;
    @api fieldLabel; 
    @track price;
    @track error;
    @track isLoading = true;
    @track weight;

    @wire(getRecord, { recordId: '$recordId', fields: ['Product2.ShippingWeight'] })
    wiredProduct({ error, data }) {
        if (data) {
            this.weight = data?.fields?.ShippingWeight?.value;
        } else if (error) {
            console.error('Erro ao buscar o peso do produto:', error);
        }
    }

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

    get precoPorKg() {
        if (this.price && this.weight && parseFloat(this.weight) > 0) {
            return this.price / parseFloat(this.weight);
        }
        return null;
    }
}