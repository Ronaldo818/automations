import { LightningElement } from 'lwc';
import getFinancialData from '@salesforce/apex/CartFinancialInfoController.getFinancialData';

export default class CartFinancialInfo extends LightningElement {
    financialData;
    error;
    errorMessage;

    connectedCallback() {
        this.fetchData();
    }

    fetchData() {
        getFinancialData()
            .then((result) => {
                this.financialData = result;
                this.error = undefined;
            })
            .catch((err) => {
                this.error = err;
                this.financialData = undefined;
                this.errorMessage = err.body ? err.body.message : err.message;
            });
    }
}