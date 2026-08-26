import { LightningElement, wire } from 'lwc';
import { CartItemsAdapter, refreshCartSummary } from 'commerce/cartApi';
import { NavigationMixin } from 'lightning/navigation';
import { resolve } from 'c/cmsResourceResolver';
import { isInSitePreview } from "c/utils";
import Toast from 'lightning/toast';
import ToastContainer from 'lightning/toastContainer';
import { registerListener, unregisterAllListeners} from 'c/pubsub';
import { CurrentPageReference } from 'lightning/navigation'; 

import LOCALE from "@salesforce/i18n/locale";
import CURRENCY from "@salesforce/i18n/currency";

/**
 * @slot region1
 * @slot region2
 */

export default class CommerceCustomCartItems extends NavigationMixin(LightningElement) {

    _items = [];
    _webstoreId;
    _accountId;
    _connectedResolver;
    _validCart = true;
    total;
    isPreview;

    _canResolveUrls = new Promise((resolved) => {
        this._connectedResolver = resolved;
    });


    get finishDisabled(){
        return (this.refreshCounter > 0);
    }

    get webstoreId() {
        return this._webstoreId;
    }

    get accountId() {
        return this._accountId;
    }

    get displayItems() {
        return this._items;
    }

    set displayItems(items) {
        const generatedUrls = [];
        this._items = (items || []).map((item) => {
            const newItem = { ...item };

            newItem.descontoPermitido = newItem.cartItem.customFields[0]?.PercentualDescontoEscalonado__c ?? 0;
            newItem.descontoNegociado = newItem.cartItem?.customFields[0]?.DescontoNegociado__c ?? 0;
            newItem.valorNegociado = (newItem.cartItem?.customFields[0]?.ValorNegociado__c ?? 0);
            newItem.valorNegociadoMaximo = (newItem.cartItem?.customFields[0]?.ValorNegociadoMaximo__c ?? 0);
            newItem.valorNegociadoST = newItem.cartItem?.customFields[0]?.PrecoKgComST__c ?? 0;
            newItem.precoKg = newItem.cartItem?.customFields[0]?.PriceKG__c ?? 0;
            newItem.pesoTotal = newItem.cartItem?.customFields[0]?.PesoTotal__c ?? 0;
            newItem.valorMinimo = newItem.cartItem?.customFields[0]?.ValorMinimo__c ?? 0;
            newItem.farol = newItem.cartItem?.customFields[0]?.Farol__c ?? 0;
            newItem.verbaConsumida = newItem.cartItem?.customFields[0]?.ValorVerba__c ?? 0;
            newItem.masterOrderItemId = newItem.cartItem?.customFields[0]?.MasterOrderItemId__c;
            newItem.productName = newItem.cartItem?.productDetails.fields.DescricaoForcaVendas__c;

            newItem.productUrl = '';
            newItem.productImageUrl = resolve(item.cartItem.productDetails.thumbnailImage.url);
            newItem.productImageUrl = this.replaceSecondOccurrence(newItem.productImageUrl, '/', '/sfsites/c/');
            newItem.productImageAlternativeText = item.cartItem.productDetails.thumbnailImage.alternateText || '';

            // Get URL for the product, which is asynchronous and can only happen after the component is connected to the DOM (NavigationMixin dependency).
            const urlGenerated = this._canResolveUrls
                .then(() =>
                    this[NavigationMixin.GenerateUrl]({
                        type: 'standard__recordPage',
                        attributes: {
                            recordId: newItem.cartItem.productId,
                            objectApiName: 'Product2',
                            actionName: 'view'
                        }
                    })
                )
                .then((url) => {
                    newItem.productUrl = url;
                });
            generatedUrls.push(urlGenerated);
            return newItem;
        });

        Promise.all(generatedUrls).then(() => {
            this._items = Array.from(this._items);
        });
    }

    @wire(CurrentPageReference) pageRef;

    connectedCallback() {
        registerListener('goCheckout', this.handleFinalizarPedido, this);

        this._connectedResolver();
        this.isPreview = isInSitePreview();

        const toastContainer = ToastContainer.instance();
        toastContainer.maxToasts = 5;
        toastContainer.toastPosition = 'top-center';
    }


    disconnectedCallback() {
        unregisterAllListeners(this);

        this._canResolveUrls = new Promise((resolved) => {
            this._connectedResolver = resolved;
        });
    }

    @wire(CartItemsAdapter, {pageSize: 100})
    cartSummary({ error, data }) {
        if (data) {
            this._accountId = data.cartSummary.accountId;
            this._webstoreId = data.cartSummary.webstoreId;
            this.displayItems = data.cartItems;
        } else if (error) {
            console.error(error);
        }
    }

    replaceSecondOccurrence(string, target, replacement) {
        let firstIndex = string.indexOf(target);
        if (firstIndex === -1) {
            return string;
        }

        let secondIndex = string.indexOf(target, firstIndex + 1);
        if (secondIndex === -1) {
            return string;
        }

        return string.substring(0, secondIndex) + replacement + string.substring(secondIndex + target.length);
    }

    refreshCounter = 0;
    handleRefresh(event){

        if(event.detail.type == 'add') this.refreshCounter++;
        if(event.detail.type == 'remove') this.refreshCounter--;

        clearTimeout(this.tempo);
        if(event.detail.type != 'clear'){
            this.tempo = setTimeout(() => {
                if(this.refreshCounter == 0){
                    refreshCartSummary().then(() => {

                    }).catch(error => {
                        console.error('handleRefresh', error);
                    });
                }                
            }, 1000);
        }
    }

    errorCounter = 0;
    handleValidity(event){
        this._validCart = event.detail.type;
    }

    handleFinalizarPedido(event){
        let isFormValid = [...this.template.querySelectorAll('c-commerce-custom-cart-items-item')]
        .reduce((validSoFar, input_Field_Reference) => {
            input_Field_Reference.checkFormValidity();
            return validSoFar && input_Field_Reference.checkFormValidity();
        }, true);

        if (isFormValid) {
            this[NavigationMixin.Navigate]({
                type: 'comm__namedPage',
                attributes: {
                    name: 'Current_Checkout'
                },
            });
        } else {
            Toast.show({
                label: 'Corrija todos os erros de preenchimento antes de continuar',
                variant: 'warning',
                mode: 'dismissible'
            }, this);
        }
    }

}