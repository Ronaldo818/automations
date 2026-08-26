import { LightningElement, api, wire } from 'lwc';
import Toast from 'lightning/toast';
import ToastContainer from 'lightning/toastContainer';
import { CartItemsAdapter } from 'commerce/cartApi';
import FORM_FACTOR from '@salesforce/client/formFactor';
import CreateMasterOrderModal from 'c/fvCreateMasterOrderModal'
import isAccountBlocked from '@salesforce/apex/CommerceUtil.isAccountBlocked';
import { effectiveAccount } from 'commerce/effectiveAccountApi';

export default class FvCreateMasterOrder extends LightningElement {

     

    isProcessing = false;
    showModal = false

    comment

    masterOrderId;
    masterOrderName;

    maxDurationDays;

    created = false;

    cartItemsValidity;


    _cartDetails
    @api
    set cartDetails(value){
        if(value){
            this._cartDetails = value;
            //console.log('cartDetails: ', value);
        }
    }
    get cartDetails(){
        return this._cartDetails
    }

    connectedCallback(){
        const toastContainer = ToastContainer.instance();
        toastContainer.maxToasts = 5;
        toastContainer.toastPosition = 'top-center';
    }

    _cartItems
    @wire(CartItemsAdapter, {pageSize: 100})
    cartSummary({ error, data }) {
        if (data) {
            console.log('CartItemsAdapter data: ', data);
            this._cartItems = data.cartItems.map((item) => ({
                productCode: item.cartItem.productDetails.sku,
                productId: item.cartItem.productDetails.productId,
                quantity: item.cartItem.quantity,
                listPrice: item.cartItem.listPrice,
                negotiatedPrice: item.cartItem.unitAdjustedPriceWithItemAdj,
                listPriceKg: item.cartItem.customFields[0]?.PriceKG__c,
                minPriceKg: item.cartItem.customFields[0]?.ValorMinimo__c,
                negotiatedPriceKg: item.cartItem.customFields[0]?.ValorNegociado__c,
                
            }));
        } else if (error) {
            console.error(error);
        }
    }

    @api
    set cartItems(value){
    }

    @api
    effectiveAccount

    get cartItems(){
        return this._cartItems;
    }

    get isButtonDisabled(){
        return this.isModalOpen ||  this.cartItems == undefined || this.cartItems.length == 0;
    }

    get modalSize(){
        //console.log(FORM_FACTOR);
        if(FORM_FACTOR === 'Small'){
            return 'large';
        }
        return 'medium'
    }

    isModalOpen = false;
    
    async handleOpenModal(){
        const effectiveAccountId = effectiveAccount.accountId;
        await isAccountBlocked({accountId: effectiveAccountId})
        .then(result => {
            if(result == true){
                Toast.show({label: 'Conta Bloqueada',
                            message: 'Não será permitido gerar Mesa de Negociação',
                            variant: 'error',
                            mode: 'dismissible'
                }, this);
            } else {
                this.isModalOpen = true;
                const optionsToSend = 
                [
                    { name: 'cartId', value: this.cartDetails.cartId},
                    { name: 'accountName', value: this.effectiveAccount},
                    { name: 'cartDetails', value: this.cartDetails},
                    { name: 'cartItems', value: this.cartItems},
                ]

                //console.log(optionsToSend);
                const result = CreateMasterOrderModal.open({
                    // `label` is not included here in this example.
                    // it is set on lightning-modal-header instead
                    label: 'info',
                    size: this.modalSize,
                    description: 'Info Modal',
                    content: 'Passed into content api',
                    options: optionsToSend
                });
                this.isModalOpen = false;
                // if modal closed with X button, promise returns result = 'undefined'
                // if modal closed with OK button, promise returns result = 'okay'
                //console.log(result);
            }

        })
        .catch(error => {
            console.error(error);
        })
    }

}