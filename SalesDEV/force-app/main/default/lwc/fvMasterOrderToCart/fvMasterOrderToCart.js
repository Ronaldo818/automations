import { LightningElement, api } from 'lwc';
import basePath from "@salesforce/community/basePath";
import { effectiveAccount } from 'commerce/effectiveAccountApi'; 
import masterOrderToCart from '@salesforce/apex/MasterOrderToCart.masterOrderToCart';

export default class FvMasterOrderToCart extends LightningElement {
    @api
    masterOrderId

    showModal = false;

    isProcessing = false;

    handleConvert(){
        this.isProcessing = true;
        masterOrderToCart({masterOrderId: this.masterOrderId, accId: effectiveAccount.accountId})
        .then((result) => {
            this.isProcessing = false;
            window.open(basePath + '/cart','_self');
        }).catch((err) => {
            this.isProcessing = false;
        });
    }

    handleOpenModal(){
        this.showModal = true;
    }

    handleCloseModal(){
        this.showModal = false;
    }
}