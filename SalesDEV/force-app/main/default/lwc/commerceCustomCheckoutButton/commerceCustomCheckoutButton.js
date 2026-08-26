import { LightningElement, wire, api } from 'lwc';
import { CartItemsAdapter } from 'commerce/cartApi';
import { CurrentPageReference } from 'lightning/navigation'; 
import { fireEvent } from 'c/pubsub';
import LOCALE from "@salesforce/i18n/locale";
import CURRENCY from "@salesforce/i18n/currency";

export default class CommerceCustomCheckoutButton extends LightningElement {

    @api label;
    @api variant;
    @api showValue;
    @api stretch;

    total;

    get labelTotal(){
        return (this.label + (this.showValue ? '(' + this.total + ')' : ''));
    }

    @wire(CurrentPageReference) pageRef;

    @wire(CartItemsAdapter, {pageSize: 100})
    cartSummary({ error, data }) {
        if (data) {
            this.total = new Intl.NumberFormat(LOCALE, {    style: "currency",
                                                            currency: CURRENCY,
                                                            currencyDisplay: "symbol", }).format(data.cartSummary?.totalProductAmountAfterAdjustments ?? 0);
        } else if (error) {
            console.error(error);
        }
    }

    handleFinalizarPedido(event){
        fireEvent(this.pageRef, 'goCheckout')
    }
}