import { LightningElement, api, wire } from 'lwc';

import FORM_FACTOR from '@salesforce/client/formFactor';


import decimalSeparator from '@salesforce/i18n/number.decimalSeparator';
import groupingSeparator from '@salesforce/i18n/number.groupingSeparator';


import addItemToCart from '@salesforce/apex/FvCartController.addItemToCart';
import { refreshCartSummary, CartSummaryAdapter } from 'commerce/cartApi';
import {
    createWishlistItemAddAction,
    dispatchAction
} from 'commerce/actionApi';
import ToastContainer from 'lightning/toastContainer';
import Toast from 'lightning/toast';

const ADD_TO_CART_LABEL = 'Adicionar ao Carrinho';
const UPDATE_CART_LABEL = 'Confirmar Alterações';




export default class FvProductPriceAndQuantitySelector extends LightningElement {

    getDecimalSeparator = () => decimalSeparator;
    getGroupingSeparator = () => groupingSeparator;

    isProcessing = true;
    _product

    cartId
    @wire(CartSummaryAdapter, {})
    cartSummary({ error, data }) {
        if (data) {
            this.cartId = data.cartId;
            //console.log('got cartId', this.cartId);
        } else if (error) {
            console.error(error);
        }
    }

    @api
    set product(value) {
        if (value) {
            this._product = value;
        }
    }

    _effectiveAccountId

    @api
    set effectiveAccountId(value){
        if(value){
            this._effectiveAccountId = value;
        }
        //console.log('set effectiveAccountId: ', value);
    }

    get effectiveAccountId(){
        return this._effectiveAccountId;
    }

    get product() {
        return this._product;
    }

    _productPricing

    @api
    set productPricing(value) {
        if (value) {
            this._productPricing = value;
            this._negotiatedPrice = value.unitPrice;
        }
    }

    get productPricing() {
        return this._productPricing;
    }

    get listPrice() {
        return this._productPricing?.unitPrice;
    }

    _quantity = 1;
    displayQuantity = '1'
    set quantity(value) {
        this._quantity = Number(value);
        this.displayQuantity = String(value);
        this.refs.quantitySelector.value = this.displayQuantity;
    }

    get quantity() {
        return this._quantity;
    }

    get isDecrementButtonDisabled() {
        return this.quantity <= 1 || this.isProcessing;
    }

    get isIncrementButtonDisabled() {
        return this.isProcessing || (this.maxQty && this.quantity >= this.maxQty);
    }

    _negotiatedPrice;
    _discountPercentage = 0;

    set negotiatedPrice(value) {
        if(value){
            this._negotiatedPrice = value;
            this._discountPercentage = 100 - value * 100 / this.productPricing.unitPrice;
        }
    }

    get negotiatedPrice() {
        return this._negotiatedPrice
    }

    set discountPercentage(value) {
        this._discountPercentage = value;
        this._negotiatedPrice = this.productPricing.unitPrice - this._discountPercentage / 100 * this.productPricing.unitPrice
    }

    get discountPercentage() {
        return this._discountPercentage
    }

    get pattern() {
        // eslint-disable-next-line no-useless-escape
        return `[+\-]?(\\d*[${this.getGroupingSeparator()}]?)*[${this.getDecimalSeparator()}]?\\d*`;
    }

    maxQty = null

    masterOrderItemId

    async connectedCallback() {
        const toastContainer = ToastContainer.instance();
        toastContainer.maxShown = 5;
        toastContainer.toastPosition = 'top';
    }

    handleAddToCart() {
        this.isProcessing = true;
        addItemToCart({ productId: this.product.id, quantity: this.quantity, negotiatedPrice: this.negotiatedPrice, masterOrderItemId: this.masterOrderItemId })
            .then((result) => {
                refreshCartSummary();
                Toast.show({
                    label: this.showAlreadyInCart ? 'Item atualizado' : 'Item adicionado ao carrinho',
                    /*message: this.showAlreadyInCart ? 'Item atualizado' : 'Item adicionado ao carrinho',*/
                    mode: 'dismissible',
                    variant: 'success'
                }, this);
                this.refs.masterOrderPicker.update();
                this.isProcessing = false;
            }).catch((err) => {
                this.isProcessing = false;
            });
    }

    handleAddToWishlist() {
        dispatchAction(this, createWishlistItemAddAction(this.product.id), {
            onSuccess: () => {
                Toast.show({
                    label: 'Produto adicionado aos favoritos',
                    /*message: 'Produto adicionado aos favoritos',*/
                    mode: 'dismissible',
                    variant: 'success'
                }, this);
            },
            onError: () => {
                if (!this.sessionContext?.data?.isLoggedIn) {
                    window.open(basePath + '/login', '_self');
                    this.inWishList = !this.inWishList;
                    this.currentProduct.inWishList = this.inWishList;
                } else {
                    Toast.show({
                        label: 'Error',
                        message: 'Ocorreu um erro ao tentar adicionar produto aos favoritos',
                        mode: 'sticky',
                        variant: 'error'
                    }, this);
                }
            },
        });
    }

    handleInputChange(event) {
        const value = event?.target.value;
        let tmp = value.replaceAll('/[^0-9]+', '');

        if (isNaN(tmp) || isNaN(parseFloat(tmp)) || 0 >= parseFloat(tmp)) {
            this.quantity = '1';
            event.target.value = '1';
        }
        else {
            if(this.maxQty && this.maxQty < parseFloat(tmp)){
                tmp = this.maxQty;
            }
            this.quantity = tmp;
        }
    }

    increment(event) {
        this.quantity = this.quantity + 1;
    }

    decrement(event) {
        this.quantity = this.quantity - 1;
    }

    get productId() {
        return this.product?.id
    }

    get mainGridStyle() {
        if (FORM_FACTOR == 'Small') {
            return 'max-width: 800px; justify-content: center;'
        }
        else {
            return 'max-width: 800px;'
        }
    }

    previousNegotiatedPrice;

    _showAlreadyInCart = false;

    addToCartLabel = ADD_TO_CART_LABEL;

    set showAlreadyInCart(value) {
        this._showAlreadyInCart = value;
        this.addToCartLabel = value ? UPDATE_CART_LABEL : ADD_TO_CART_LABEL;
    }

    get showAlreadyInCart() {
        return this._showAlreadyInCart
    }

    async handleMasterOrderSelected(event) {
        this.masterOrderItemId = event.detail.value === 'null' ? null : event.detail.value;
        if (event.detail.init && event.detail.qty !== 0) {
            this.quantity = event.detail.qty;
            this.showAlreadyInCart = true;
        }
        if (this.masterOrderItemId) {
            this.maxQty = event.detail.maxQty;
            if (this.quantity > this.maxQty) {
                this.quantity = this.maxQty
            }
            this.previousNegotiatedPrice = this._negotiatedPrice;
            this.negotiatedPrice = event.detail.nPrice;
        }
        else {
            this.negotiatedPrice = this.previousNegotiatedPrice || event.detail.nPrice;
            this.maxQty = null
        }
        this.isProcessing = false;
    }
}