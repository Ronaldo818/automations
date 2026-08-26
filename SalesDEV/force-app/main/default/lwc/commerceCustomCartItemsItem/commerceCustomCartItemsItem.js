import { LightningElement, api } from 'lwc';
import { deleteItemFromCart } from 'commerce/cartApi';
import updateCartItem from '@salesforce/apex/CommerceCustomCartItemsController.updateCartItem';

const MAX_POSSIBLE_VALUE = 999999

export default class CommerceCustomCartItemsItem extends LightningElement {

    lineItem;
    quantity;
    price;
    discount;
    _webstoreId = '';
    _accountId = '';
    productId = '';
    masterOrderItemId;
    maxQty = MAX_POSSIBLE_VALUE;
    maxPrice;
    productName;
    isLoading = false;

    @api
    get cartItem(){
        return this.lineItem;
    }

    set cartItem(value){
        if(value){
            this.lineItem = value;
            this.quantity = parseInt(value.cartItem.quantity, 10); //value.cartItem.quantity;
            this.productId = value.cartItem.productId;
            this.price = value.valorNegociado;
            this.maxPrice = value.valorNegociadoMaximo; // Valor máximo negociável de até 100% no preço de lista
            this.discount = value.descontoNegociado;
            this.masterOrderItemId = value.masterOrderItemId;
            this.productName = value.productName;
            this.loaded = true;
        }
    }

    @api
    get webstoreId(){
        return this._webstoreId;
    }
    set webstoreId(value){
        if(value){
            this._webstoreId = value;
        }
    }

    @api
    get accountId(){
        return this._accountId;
    }
    set accountId(value){
        if(value){
            this._accountId = value;
        }
    }

    get disableDiscount(){
        return (this.isLoading || this.masterOrderItemId != null);
    }

    get disablePrice(){
        return (this.isLoading || this.masterOrderItemId != null);
    }

    get disableQuantity(){
        return (this.isLoading);
    }

    get isLoaded(){
        return (this.loaded && this._accountId && this._accountId != '' && this.productId && this.productId != '');
    }

    get isIncrementButtonDisabled(){
        return (this.isLoading || this.quantity >= this.maxQty);
    }

    get isDecrementButtonDisabled(){
        return (this.isLoading || this.quantity == 1);
    }   

    increment(event) {
        if(this.refs.inputQuantity && !this.isLoading){
            let qtd = parseInt(this.refs.inputQuantity.value);
            qtd++;
            this.refs.inputQuantity.value = qtd;
            this.handleItemChange();
        }
    }

    decrement(event) {
        if(this.refs.inputQuantity && !this.isLoading){
            let qtd = parseInt(this.refs.inputQuantity.value);
            if(qtd > 1){
                qtd--;
            }
            this.refs.inputQuantity.value = qtd;
            this.handleItemChange();
        }
    }

    handleClickDelete(){
        let cartItemId = this.lineItem.cartItem.cartItemId;
        deleteItemFromCart(
            cartItemId
        ).then((result) => {

        }).catch((error) => {
            console.error('handleClickDelete - deleteItemFromCart: ', error)
        });
    }

    async handleMasterOrderSelected(event) {
        let eventMasterOrderItemId = (event.detail.value === 'null' ? null : event.detail.value);

        if(eventMasterOrderItemId){
            this.maxQty = event.detail.maxQty;
        }

        if(this.masterOrderItemId != eventMasterOrderItemId){
            if(eventMasterOrderItemId == null){
                // this.quantity = 1;
                this.maxQty = MAX_POSSIBLE_VALUE;
            } else {
                this.refs.inputDiscount.value = 0;
                this.refs.inputPrice.value = event.detail.nPrice;
                this.refs.inputPrice.reportValidity();
                this.refs.inputDiscount.reportValidity();
                
                if (this.quantity > this.maxQty) {
                    this.quantity = this.maxQty
                }
            }

            this.refs.inputQuantity.reportValidity();
            this.masterOrderItemId = eventMasterOrderItemId;
            this.updateItem(true);
        }
    }

    handleDiscountChange(event){
        clearTimeout(this.discountTime);
        this.discountTime = setTimeout(() => {
            if(this.dicount != this.refs.inputDiscount.value){
                let desc = this.lineItem.precoKg - (this.refs.inputDiscount.value / 100 * this.lineItem.precoKg)
                this.refs.inputPrice.value = desc.toFixed(2);
                this.updateItem(false);
            }
        }, 200);
    }

    handleItemChange(event){
        this.updateItem(false);
    }

    handleBlur(event){
        this.updateItem(false);
    }

    @api
    checkFormValidity(){
        return [...this.template.querySelectorAll('lightning-input')]
        .reduce((validSoFar, input_Field_Reference) => {
            input_Field_Reference.reportValidity();
            return validSoFar && input_Field_Reference.reportValidity();
        }, true);
    }

    async updateItem(forceUpdate){

        clearTimeout(this.globalTime);
        const clearEvent = new CustomEvent("refresh", { detail: { type: 'clear'} });
        this.dispatchEvent(clearEvent);

        if(this.checkFormValidity()){
            if(this.price != this.refs.inputPrice.value || this.quantity != this.refs.inputQuantity.value || forceUpdate){
                this.globalTime = setTimeout(() => {

                    if(this.masterOrderItemId != null && this.refs.inputQuantity.value > this.maxQty){
                        this.refs.inputQuantity.value = this.maxQty;
                    }

                    this.isLoading = true;
                    const selectedEvent = new CustomEvent("refresh", { detail: { type: 'add'} });
                    this.dispatchEvent(selectedEvent);

                    updateCartItem({
                        webstoreId: this._webstoreId,
                        effectiveAccountId: this._accountId,
                        activeCartOrId: this.lineItem.cartItem.cartId,
                        cartItemId: this.lineItem.cartItem.cartItemId,
                        quantity: this.refs.inputQuantity.value,
                        price: this.refs.inputPrice.value,
                        masterOrderItemId: this.masterOrderItemId
                    }).then(result => {
                        const refreshEvent = new CustomEvent("refresh", { detail: { type: 'remove'} });
                        this.dispatchEvent(refreshEvent);
                    }).catch(error => {
                        console.error('handleItemChange - updateCartItem: ', error);
                    }).finally(f =>{
                        this.isLoading = false;
                    })
                }, 1500);
            }
        }
    }

}