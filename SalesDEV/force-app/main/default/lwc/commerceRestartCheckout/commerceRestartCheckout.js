import { LightningElement, wire } from 'lwc';
import { CurrentPageReference } from "lightning/navigation";

import recalculateTaxes from '@salesforce/apex/B2BCheckoutController.recalculateTaxes';



export default class CommerceRestartCheckout extends LightningElement {

    @wire(CurrentPageReference)
    pageRef

    get inBuilder(){
        return this.pageRef.state.view === "editor";
    }

    connectedCallback(){
        console.log('Requesting tax recalc')
        recalculateTaxes();
    }
}