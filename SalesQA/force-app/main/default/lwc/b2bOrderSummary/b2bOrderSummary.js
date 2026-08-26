import { LightningElement, api, wire, track } from 'lwc';

export default class B2bOrderSummary extends LightningElement {
    @api title;
    @track __subtotal = 0;
    @track __promocao = 0;
    @track __impostos = 0;
    @track __cashBack = 0;

}