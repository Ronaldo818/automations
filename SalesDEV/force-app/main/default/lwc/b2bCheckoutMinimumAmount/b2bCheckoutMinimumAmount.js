import { wire } from 'lwc';
import { CheckoutComponentBase, CheckoutInformationAdapter} from 'commerce/checkoutApi';
import { CartSummaryAdapter } from 'commerce/cartApi';
import { CurrentPageReference } from "lightning/navigation";

import getLimits from '@salesforce/apex/B2BCheckoutMinimumAmountController.getLimits'

const CheckoutStage = {
    CHECK_VALIDITY_UPDATE: 'CHECK_VALIDITY_UPDATE',
    REPORT_VALIDITY_SAVE: 'REPORT_VALIDITY_SAVE',
    BEFORE_PAYMENT: 'BEFORE_PAYMENT',
    PAYMENT: 'PAYMENT',
    BEFORE_PLACE_ORDER: 'BEFORE_PLACE_ORDER',
    PLACE_ORDER: 'PLACE_ORDER'
};

export default class B2bCheckoutMinimumAmount extends CheckoutComponentBase {

    @wire(CurrentPageReference)
    pageRef

    isLoading = true;
    isValid  = false;
    isValidMaxPack  = false;
    minimumAmount = 0;
    maxQuantity = 0;
    totalCustomerMonthPacks = 0;
    pessoaFisica;

    get isInBuilder(){
        return this.pageRef.state.view === "editor";
    }

    get iconName(){
        return this.isValid ? 'utility:success' : 'utility:error'
    }

    get iconVariant(){
        return this.isValid ? 'success' : 'error'
    }

    get iconNameMaxPack(){
        return this.isValidMaxPack ? 'utility:success' : 'utility:error'
    }

    get iconVariantMaxPack(){
        return this.isValidMaxPack ? 'success' : 'error'
    }

    get quantityLimit(){
        return ((this.maxQuantity > 0) && this.pessoaFisica);
    }

    checkoutStatus;
    pendingDispatchCommit = false;

    @wire(CheckoutInformationAdapter, {})
    checkoutInformation({ error, data }) {
        if (data) {
            this.checkoutStatus = data.checkoutStatus
            if(this.checkoutStatus == 200 && this.pendingDispatchCommit){
                this.dispatchCommit();
            }
            // console.log('CheckoutInformationAdapter:', data);
        } else if (error) {
            console.error('CheckoutInformationAdapter:', error);
        }
    }

    @wire(CartSummaryAdapter, {})
    cartSummary({ error, data }) {
        if (data) {
            getLimits({
                accountId: data.accountId
            })
            .then((result) => {
                const minimumAmount = result.minimumAmount;
                const cartAmount = parseFloat(data.grandTotalAmount)
                this.isValid =  cartAmount >= minimumAmount;
                this.minimumAmount = result.minimumAmount;

                this.pessoaFisica = data.customFields[0]?.NaturezaJuridicaCliente__c == 'PF' ?? false;
                let maxPackPerMonth = parseFloat(result.maxPackPerMonth);
                this.totalCustomerMonthPacks = parseFloat(result.totalCustomerMonthPacks) + parseFloat(data.totalProductCount);
                this.maxQuantity = result.maxPackPerMonth;
                this.isValidMaxPack = !this.quantityLimit || (this.totalCustomerMonthPacks < maxPackPerMonth);

                this.isLoading = false;
                if(this.checkoutStatus == 200){
                    this.dispatchCommit();
                }
                else{
                    this.pendingDispatchCommit = true;
                }
            }).catch((err) => {
                console.error('B2bCheckoutMinimumAmount - CartSummaryAdapter:', err);
            });
            
        } else if (error) {
            console.error(error);
        }
    }

    reportValidity(){
        // console.log('reporting validity');
        if(this.isValid && this.isValidMaxPack){
            /*
            this.dispatchUpdateErrorAsync({
                groupId: "MinimumAmount",
            });
            */
            this.dispatchUpdateAsync({
                notifications:[{
                    groupId: "MinimumAmount",
                    //type: "/commerce/errors/checkout-failure",
                    //detail: "Valor mínimo não atendido",
                }]
            })
        }
        else{
            /*
            this.dispatchUpdateErrorAsync({
                groupId: "MinimumAmount",
                type: "/commerce/errors/checkout-failure",
                exception: "Valor mínimo não atendido",
            });
            */
           if(!this.isValid){
                this.dispatchUpdateAsync({
                    notifications:[{
                        groupId: "MinimumAmount",
                        type: "/commerce/errors/checkout-failure",
                        detail: "Valor mínimo não atendido",
                    }]
                })
           } else if (!this.isValidMaxPack){
                this.dispatchUpdateAsync({
                    notifications:[{
                        groupId: "MinimumAmount",
                        type: "/commerce/errors/checkout-failure",
                        detail: "Excedido limite de caixas por mês. Total incluindo este pedido: " + this.totalCustomerMonthPacks
                    }]
                })
           }
        }
        return this.isValid && this.isValidMaxPack;
    }

    async stageAction(checkoutStage) {
        // console.log('B2bCheckoutMinimumAmount - ', checkoutStage);
        switch (checkoutStage) {
            case CheckoutStage.REPORT_VALIDITY_SAVE:
                return await Promise.resolve(this.reportValidity());
            default:
                return Promise.resolve(true);
        }
    }
}