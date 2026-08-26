import { LightningElement, wire } from 'lwc';
import { CartSummaryAdapter, refreshCartSummary } from 'commerce/cartApi'
import getTableData from '@salesforce/apex/fvCartItemsEditorController.getTableData';
import batchAddToCart from '@salesforce/apex/fvCartItemsEditorController.batchAddToCart';
import FORM_FACTOR from '@salesforce/client/formFactor';


const cart_items_collumns = [
    { label: 'Código', fieldName: 'pCode', type: 'text', initialWidth: 100},
    { label: 'Nome', fieldName: 'pName', type: 'text', initialWidth: 200},
    { label: 'Tabela', fieldName: 'lPrice', type: 'currency'},
    { label: 'Negociado', fieldName: 'nPrice', type: 'currency'},
    { label: 'Disp.', fieldName: 'qtyA', type: 'number'},
    { label: 'Qtd.', fieldName: 'qty', type: 'number', editable: true, typeAttributes: { maximumFractionDigits: 0, step: '1' } },
    { label: 'Total', fieldName: 'total', type: 'currency'}
];

const cart_items_collumns_small = [
    { label: 'Código', fieldName: 'pCode', type: 'text', initialWidth: 100},
    { label: 'Negociado', fieldName: 'nPrice', type: 'currency'},
    { label: 'Disp.', fieldName: 'qtyA', type: 'number'},
    { label: 'Qtd.', fieldName: 'qty', type: 'number', editable: true, typeAttributes: { maximumFractionDigits: 0, step: '1' } },
    { label: 'Total', fieldName: 'total', type: 'currency'}
];

export default class FvCartItemsEditor extends LightningElement {

    masterOrderId
   

    cartItemsCollumns = FORM_FACTOR === 'Small' ? cart_items_collumns_small : cart_items_collumns;

    cartItemsData = []

    isProcessing = false
    isLoading = true;

    @wire(CartSummaryAdapter, {})
    cartInfo({ error, data }) {
        if (data) {
            console.log('FvCartItemsEditor -  CartSummaryAdapter:', data);
                this.masterOrderId = data.customFields[0]?.MasterOrderId__c;
                    getTableData({masterOrderId: this.masterOrderId})
                    .then((result) => {
                        this.cartItemsData = JSON.parse(result);
                        this.isLoading = false;
                        this.isProcessing = false;
                        this.refs.cartItemsDataTable.draftValues = [];
                        console.log('this.cartItemsData:', this.cartItemsData);
                    }).catch((err) => {
                        console.error(err);
                        this.isLoading = false;
                    });
        } else if (error) {
            console.error("CommerceCartIsProcessing cartInfo: ", error);
        }
    }

    handleInlineEdit(event){
        console.log(event);
        console.log(event.detail);
    }

    handleSave(event){
        console.log(event);
        console.log(event.detail);
        this.isProcessing = true;
        batchAddToCart({lstProductIdQty: event.detail.draftValues})
        .then((result) => {
            refreshCartSummary();
        }).catch((err) => {
            console.error(err);
            this.isProcessing = false;
        });
    }
}