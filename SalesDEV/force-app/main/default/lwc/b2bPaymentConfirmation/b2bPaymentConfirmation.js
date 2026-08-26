import { LightningElement, track } from 'lwc';

import ToastContainer from 'lightning/toastContainer';
import Toast from 'lightning/toast';

import getPaymentData from '@salesforce/apex/B2BPaymentConfirmationController.getPaymentData';

export default class B2bPaymentConfirmation extends LightningElement {

    boleto = false;
    boletoLink;
    numbers;
    image;

    @track
    paymentData = {};

    toastContainer

    connectedCallback() {
 
        this.parameters = this.getQueryParameters();
        this.image = this.parameters['link'];
        this.boletoLink = this.parameters['boleto'];
        this.numbers = this.parameters['numbers'];
        console.log(this.boletoLink);
        if(this.boletoLink != null && this.boletoLink != ''){
            this.boleto = true;
        }
        console.log(this.boleto);

        getPaymentData({orderNumber: this.parameters['orderNumber']})
        .then((result) => {
            this.paymentData = JSON.parse(result);
            console.log(this.paymentData);
        }).catch((err) => {
            this.retryGetData();
        });

        const toastContainer = ToastContainer.instance();
        toastContainer.maxShown = 5;
        toastContainer.toastPosition = 'top';
    }

    retryGetData(){
        setTimeout(() => {
            getPaymentData({orderNumber: this.parameters['orderNumber']})
            .then((result) => {
                this.paymentData = JSON.parse(result);
                console.log(this.paymentData);
            }).catch((err) => {
                this.retryGetData();
            });
        }, 1000);
    }

    abrirBoleto() {
        console.log(this.boletoLink);
        window.open(this.boletoLink);
    }

    getQueryParameters() {

        var params = {};
        var search = location.search.substring(1);

        if (search) {
            params = JSON.parse('{"' + search.replace(/&/g, '","').replace(/=/g, '":"') + '"}', (key, value) => {
                return key === "" ? value : decodeURIComponent(value)
            });
        }

        return params;
    }


    copyToClipboard(event){
        navigator.clipboard.writeText(this.paymentData.copyAndPaste);

        Toast.show({
            label: 'Copiado!',
            message: 'O código Pix foi copiado para sua área de tranferência',
            mode: 'dismissible',
            variant: 'success'
        }, this);
    }

}