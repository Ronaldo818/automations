import { wire } from 'lwc';
import { CheckoutComponentBase, CheckoutInformationAdapter} from 'commerce/checkoutApi';
import { CartItemsAdapter } from 'commerce/cartApi';
import { CurrentPageReference } from "lightning/navigation";
import basePath from '@salesforce/community/basePath';


import checkCartItems from '@salesforce/apex/B2BCheckoutMinimumAmountController.checkCartItems';


const CheckoutStage = {
    CHECK_VALIDITY_UPDATE: 'CHECK_VALIDITY_UPDATE',
    REPORT_VALIDITY_SAVE: 'REPORT_VALIDITY_SAVE',
    BEFORE_PAYMENT: 'BEFORE_PAYMENT',
    PAYMENT: 'PAYMENT',
    BEFORE_PLACE_ORDER: 'BEFORE_PLACE_ORDER',
    PLACE_ORDER: 'PLACE_ORDER'
};

export default class B2bCartItemsValidations extends CheckoutComponentBase {
    
    get isValid(){
        return this.invalidCartItems.length == 0;
    }

    get showComponent(){
        return !this.isValid || this.isInBuilder;
    }

    pendingDispatchCommit = false;

    @wire(CurrentPageReference)
    pageRef

    invalidCartItems = []

    cartItemIdXInfo = {}

    get isInBuilder(){
        return this.pageRef.state.view === "editor";
    }

    @wire(CartItemsAdapter, {})
    cartItemsSummary({ error, data }){
        if(data){
            console.log('cartItemsSummary data:', data);
            let cartItemsToCheckMasterOrderItemQuantity = []
            for(let i = 0; i < data.cartItems.length; i++){
                if(data.cartItems[i]?.cartItem?.customFields[0]?.MasterOrderItemId__c){
                    cartItemsToCheckMasterOrderItemQuantity.push(data.cartItems[i]?.cartItem?.cartItemId);
                }
                this.cartItemIdXInfo[data.cartItems[i]?.cartItem?.cartItemId] = {sku: data.cartItems[i]?.cartItem?.productDetails?.sku, name: data.cartItems[i]?.cartItem?.name, productId: data.cartItems[i]?.cartItem?.productId}
            }
            if(cartItemsToCheckMasterOrderItemQuantity.length > 0){
                checkCartItems({
                    cartItemsIds: cartItemsToCheckMasterOrderItemQuantity
                })
                .then((result) => {
                    this.invalidCartItems = result.map( elem =>{
                        return {
                            ...elem,
                            sku: this.cartItemIdXInfo[elem.cartItemId].sku,
                            name: this.cartItemIdXInfo[elem.cartItemId].name,
                            url: basePath + '/product/' + this.cartItemIdXInfo[elem.cartItemId].productId
                        }
                    });
                }).catch((err) => {
                    console.error('B2bCartItemsValidations - cartItemsSummary:', err);
                });
            }
        }
        else if(error){
            console.error(error);
        }
    }

    @wire(CheckoutInformationAdapter, {})
    checkoutInformation({ error, data }) {
        if (data) {
            this.checkoutStatus = data.checkoutStatus
            if(this.checkoutStatus == 200 && this.pendingDispatchCommit){
                this.dispatchCommit();
            }
            console.log('CheckoutInformationAdapter:', data);
        } else if (error) {
            console.error('CheckoutInformationAdapter:', error);
        }
    }

    reportValidity(){

        if(this.isValid){
            this.dispatchUpdateAsync({
                notifications:[{
                    groupId: "CartItemsValidity",
                }]
            })
        }
        else{
            this.dispatchUpdateAsync({
                notifications:[{
                    groupId: "CartItemsValidity",
                    type: "/commerce/errors/checkout-failure",
                    detail: "Carrinho possiu items inválidos",
                }]
            })
        }
        return this.isValid;
    }

    async stageAction(checkoutStage) {
        switch (checkoutStage) {
            case CheckoutStage.REPORT_VALIDITY_SAVE:
                return await Promise.resolve(this.reportValidity());
            default:
                return Promise.resolve(true);
        }
    }
}