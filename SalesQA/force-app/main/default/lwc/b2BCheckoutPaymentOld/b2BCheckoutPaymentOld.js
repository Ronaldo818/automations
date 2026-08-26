import { wire } from 'lwc';
import { CheckoutComponentBase, placeOrder, CheckoutInformationAdapter } from 'commerce/checkoutApi';
import { CartSummaryAdapter } from 'commerce/cartApi';
import { NavigationMixin } from 'lightning/navigation';

import getCashback from '@salesforce/apex/B2BPaymentController.getCashback';

import getPagBankPublicKey  from '@salesforce/apex/PagBankPublicKey.getPublicKey';

import setCashBackValue from '@salesforce/apex/B2BPaymentController.setCashBackValue'
import pay from '@salesforce/apex/B2BPaymentController.pay';

const CheckoutStage = {
    PERSISTED: 'PERSISTED',
    CHECK_VALIDITY_UPDATE: 'CHECK_VALIDITY_UPDATE',
    REPORT_VALIDITY_SAVE: 'REPORT_VALIDITY_SAVE',
    BEFORE_PAYMENT: 'BEFORE_PAYMENT',
    PAYMENT: 'PAYMENT',
    BEFORE_PLACE_ORDER: 'BEFORE_PLACE_ORDER',
    PLACE_ORDER: 'PLACE_ORDER'
};

export default class b2BCheckoutPaymentOld extends NavigationMixin(CheckoutComponentBase) {
    
    cartId
    checkoutStatus;
    showAlert = false;
    alertMessage
    alertLogId
    paymentDisabled = true;

    orderConfirmationUrlParameters = ''

    showError = false;
    error

    isLoading = false;

    paymentOptions = [
        { label: 'Boleto', value: 'boleto' },
        { label: 'Cartão de Crédito', value: 'cc' },
        { label: 'Pix', value: 'pix' }
    ];
    selectedPaymentOption = '';


    @wire(CartSummaryAdapter, {})
    cartSummary({ error, data }) {
        if (data) {
            this.cartId = data.cartId;
        } else if (error) {
            console.error(error);
        }
    }
    

    @wire(CheckoutInformationAdapter, {})
    checkoutInformation({ error, data }) {
        if (data) {
            this.checkoutStatus = data.checkoutStatus
            console.log('CheckoutInformationAdapter:', data);
        } else if (error) {
            console.error('CheckoutInformationAdapter:', error);
        }
    }

    get expiryMonths() {
        const expiryMonths = [],
            noOfMonths = 12;
        for (let month = 1; month <= noOfMonths; month++) {
            expiryMonths.push({ label: month, value: month.toString() });
        }
        return expiryMonths;
    }

    get expiryYears() {
        const expiryYears = [],
            noOfYears = 20;
        let year, i;
        for (
            year = new Date().getFullYear(), i = 0;
            i < noOfYears;
            year++, i++
        ) {
            expiryYears.push({ label: year, value: year.toString() });
        }
        return expiryYears;
    }

    get optionsParcelas() {
        return [
            { label: '1x (Sem juros)', value: '1x' },
            // { label: 'Parcelado 1x p/ 30 dias', value: '2' },
            // { label: '3', value: '3' },
        ];
    }
    selectedParcelasOption = '';

    handleParcelas(event){
        this.selectedParcelasOption = event.target.value;
    }

    boleto = false;

    cc = false;
    ccErrorMessage = '';
    cardHolderNameRequired = false;
    cardTypeRequired = false;
    expiryMonthRequired = false;
    expiryYearRequired = false;
    cvvRequired = false;
    hideCardHolderName = false;
    hideCardType = false;
    hideCvv = false;
    hideExpiryMonth = false;
    hideExpiryYear = false;
    cardHolderName;
    cardNumber;
    cardType;
    cvv;
    expiryMonth;
    expiryYear;
    email;
    telefone;

    get allCardDataIsFilled() {
        return this.cardHolderName && this.cardNumber && this.cvv && this.expiryMonth && this.expiryYear && this.selectedParcelasOption;
    }

    pix = false;

    /* Different */
    saldoCashBack;
    valorCashBack = 0;

    get cashBackLimit(){
        return Math.max(0,this.saldoCashBack);
    }

    cashbackErrorMessage = '';

    get cardNumberClass() {
        const sldsColumnSize = this.hideCvv
            ? 'slds-size_1-of-1'
            : 'slds-size_2-of-3';
        return 'slds-form-element ' + sldsColumnSize;
    }

    get disablePayment() {return this.showAlert || this.isLoading || this.selectedPaymentOption == '' || (this.selectedPaymentOption === 'cc' && !this.allCardDataIsFilled) && this.checkoutStatus === 200}


    connectedCallback() {
        // this.getPublicKey();
        this.getCustomerCashback();
    }

        async getCustomerCashback() {
        await getCashback().then(result => {
            const retorno = JSON.parse(result);
            if(retorno.error){
                this.alertLogId     = retorno.logId;
                this.alertMessage   = retorno.errorMessage;
                this.showAlert      = true;
                this.handleShowToast('Erro.', retorno.errorMessage, 'error');
                this.isLoading      = false;
            }
            else {
                if(retorno.cashbackBalance != null && retorno.cashbackBalance != undefined) {
                    const cashbakBalance = parseFloat(retorno.cashbackBalance.replace(",", "."));
                    if(cashbakBalance != 0){
                        this.saldoCashBack = retorno.cashbackBalance;
                    }
                }
            }
        })
        .catch(error => {
            this.alertMessage   = error;
            this.showAlert      = true;
            this.isLoading      = false;
        });
    }

    hasIncompleteCardPaymentFields() {
        const fieldRequiredValueMap = [
            {
                isHiddenAttr: this.hideCardHolderName,
                isRequiredAttr: this.cardHolderNameRequired,
                value: this.cardHolderName,
                error: "Nome inválido."
            },
            {
                isHiddenAttr: false,
                isRequiredAttr: true,
                value: this.cardNumber,
                error: "Número de cartão inválido."
            },
            {
                isHiddenAttr: this.hideCvv,
                isRequiredAttr: this.cvvRequired,
                value: this.cvv,
                error: "Número do CVV  inválido."
            },
            {
                isHiddenAttr: this.hideExpiryMonth,
                isRequiredAttr: this.expiryMonthRequired,
                value: this.expiryMonth,
                error:  "Mês de expiração inválido."
            },
            {
                isHiddenAttr: this.hideExpiryYear,
                isRequiredAttr: this.expiryYearRequired,
                value: this.expiryYear,
                error: "Ano de expiração inválido"
            }
        ];
        for(let i=0;  i<5;i++){
            if(fieldRequiredValueMap[i].value == undefined){
                this.ccErrorMessage = fieldRequiredValueMap[i].error;
                return true;
            }
        }

        return false;
    }

    persisted(){
        this.paymentDisabled = false;
        return true;
    }

    reportValidity() {
        if(parseFloat(this.valorCashback) != 0){
            if(parseFloat(this.valorCashback) < 0 || parseFloat(this.valorCashback) > parseFloat(this.saldoCashBack)) {
                this.cashbackErrorMessage = "Valor inválio de cashback.";
                return false;
            }
        }
        return true;
    }

    beforePayment(){
        if(this.cc){
            const incompleteFields = this.hasIncompleteCardPaymentFields();

            const componentsToValidate = this.template.querySelectorAll(
                '[data-validate]'
            );

            const validateFields = [...componentsToValidate].reduce(
                (result, component) => component.reportValidity() && result,
                true
            );

            return !incompleteFields && validateFields;
        }else if(this.pix){ 
            return true;
        }else if(this.boleto){
            return true;
        }else{
            return false;
        }
    }

    async encryptCard(cardData) {
        this.isLoading = true;
        try {
            const result = JSON.parse(await getPagBankPublicKey());
            const card = PagSeguro.encryptCard({
                publicKey: result.publicKey,
                holder: cardData.nomeCartao,
                number: cardData.numeroCartao,
                expMonth: cardData.mesCartao,
                expYear: cardData.anoCartao,
                securityCode: cardData.cvvCartao
            });
    
            const encrypted = card.encryptedCard;
            const hasErrors = card.hasErrors;
            const errors = card.errors;
            if (hasErrors) {
                const invalidKeyError = errors.find(error => error.code === 'INVALID_PUBLIC_KEY');
                if (invalidKeyError) {
                    this.alertMessage = invalidKeyError.message;
                    this.showAlert = true;
                } else {
                    this.alertMessage = JSON.stringify(errors);
                    this.showAlert = true;
                }
                return null;
            } else {
                return encrypted;
            }
        } catch (error) {
            this.alertMessage = error;
            this.showAlert = true;
            return null;
        } finally {
            this.isLoading = false;
        }
    }


    async handlePaymentType(event){
        this.selectedPaymentOption = event.detail.value;
        this.paymentOptions.forEach(elem => {
            this[elem.value] = false;
        })
        this[this.selectedPaymentOption] = true;
    }

    defaultHandleInputChange(event){
        this[event.target.dataset.fieldName] = event.target.value;
        event.target.reportValidity();
    }

    preventSensitiveInformationPropagation(keyboardEvent) {
        keyboardEvent.stopPropagation();
    }

    closeAlert() {
        this.showAlert = false;
        this.alertMessage = '';
        this.isLoading = false;
    }

    orderConfirmationParameters = {};

    async payOrder(parameter) {
        this.isLoading = true;
        await pay({jsonParameter: JSON.stringify(parameter)})
        .then(result => {
            const retorno = JSON.parse(result);
            if(retorno.error) {
                this.alertLogId     = retorno.logId;
                this.alertMessage   = retorno.errorMessage;
                this.showAlert      = true;
                this.isLoading      = false;
                // this.handleShowToast('Erro nos dados de pagamento.', 'Por favor verifique os dados digitados.', 'error');
            }
            this.orderConfirmationUrlParameters = retorno.link == undefined || retorno.link == null || retorno.link == 'null' ? '' : retorno.link;
            this.operacao   = 3;
            this.orderConfirmationParameters = {}
            if(this.orderConfirmationUrlParameters){
                const tmpList = this.orderConfirmationUrlParameters.split('&');
                tmpList.forEach(element => {
                    const tmp = element.split('=');
                    if(tmp[0]){
                        this.orderConfirmationParameters[tmp[0]] = tmp.length == 2 ? tmp[1] : '';
                    }
                });
            }
            this.isLoading  = false;
        })
        .catch(error => {
            this.alertMessage   = JSON.stringify(error);
            this.showAlert      = true;
            this.isLoading      = false;
        });
    }

    async placeOrder(){
        this.isLoading = true;
        const setCashBackResponse = JSON.parse(await setCashBackValue({valorCashback: this.refs.cashBackUsed.value }))

        if(setCashBackResponse?.error){
                this.alertLogId     = setCashBackResponse.logId;
                this.alertMessage   = setCashBackResponse.errorMessage;
                this.showAlert      = true;
                this.isLoading      = false;
            }
        const parameter = {
            formaPagamento: this.selectedPaymentOption,
            cartId: this.cartId
        }

        if(parameter.formaPagamento.toLowerCase() === 'cc' && !this.showAlert) {
            const cardParameter = {
                parcelas:       this.selectedParcelasOption,
                cvvCartao:      this.cvv,
                anoCartao:      this.expiryYear,
                mesCartao:      this.expiryMonth,
                numeroCartao:   this.cardNumber,
                nomeCartao:     this.cardHolderName  
            };
            const encryptedCard = await this.encryptCard(cardParameter);
            parameter.parcelas  = cardParameter.parcelas;
            parameter.encryptedCard = encryptedCard;
            
        }
        if(!this.showAlert){
            await this.payOrder(parameter);
        }
        if(!this.showAlert){
                    placeOrder()
        .then(result => {
            this.orderConfirmationParameters['orderNumber'] = result.orderReferenceNumber;

            this[NavigationMixin.Navigate]({
            type: 'comm__namedPage',
            attributes: {
                name: 'Order',
            },
                state: this.orderConfirmationParameters
            });
            
        }).catch(error => {
                this.alertLogId     = 
                this.alertMessage   = String.valueOf(error);
                this.showAlert      = true;
                this.isLoading      = false;
        });
        }
    }

    async stageAction(checkoutStage) {
        console.log('b2BCheckoutPayment - ', checkoutStage);
        switch (checkoutStage) {
            case CheckoutStage.PERSISTED:
                return await Promise.resolve(this.persisted());
            case CheckoutStage.REPORT_VALIDITY_SAVE:
                return await Promise.resolve(this.reportValidity());
            case CheckoutStage.BEFORE_PAYMENT:
                return await Promise.resolve(this.beforePayment());
            case CheckoutStage.PAYMENT:
                return await Promise.resolve(this.placeOrder());
            default:
                return Promise.resolve(true);
        }
    }
}