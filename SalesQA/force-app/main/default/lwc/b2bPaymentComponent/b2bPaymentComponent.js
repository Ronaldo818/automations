import { LightningElement, track, wire, api } from 'lwc';
import getPagBankPublicKey  from '@salesforce/apex/PagBankPublicKey.getPublicKey';
import changePagBankPublicKey  from '@salesforce/apex/PagBankPublicKey.changePublicKey';
import createOrder          from '@salesforce/apex/B2BLWRCheckoutControler.createOrder';
import { ShowToastEvent }   from 'lightning/platformShowToastEvent';
import pay                  from '@salesforce/apex/B2BLWRCheckoutControler.pay';
import closeWebCart         from '@salesforce/apex/B2BLWRCheckoutControler.closeWebCart';
import setPedido            from '@salesforce/apex/B2BAgrosysPedido.setPedido';
import {getRecord} from 'lightning/uiRecordApi';
import PROFILE_NAME_FIELD from '@salesforce/schema/User.Profile.Name';
import strUserId from '@salesforce/user/Id';

import basePath from '@salesforce/community/basePath';

import {subscribe, unsubscribe, APPLICATION_SCOPE, MessageContext} from 'lightning/messageService';
import checkoutMessage from '@salesforce/messageChannel/checkoutMessage__c';
import getCashback from '@salesforce/apex/B2BLWRCheckoutControler.getCashback';

const cardTypesObj = {
    Visa: 'Visa',
    'Master Card': 'Master Card',
    'American Express': 'American Express',
    'Diners Club':'Diners Club',
    Jcb: 'JCB'
};

export default class B2bPaymentComponent extends LightningElement {
    @track showAlert = false;
    @track alertMessage;
    @track alertLogId;
    @track prfName;
    @api recordId;

    _parcelas = 1;

    get optionsParcelas() {
        return [
            { label: '1x (Sem juros)', value: '1x' },
            // { label: 'Parcelado 1x p/ 30 dias', value: '2' },
            // { label: '3', value: '3' },
        ];
    }

    get parcelas(){
        return this._parcelas;
    }

    options = [
        { label: 'Boleto', value: 'boleto' },
        { label: 'Cartão de Crédito', value: 'cartao de credito' },
        { label: 'Pix', value: 'pix' }
    ];
    value = '';

    _boleto = false;
    _pix = false;
    _credit = false;
    _formaPagamento;

    _pixErrorMessage;
    _creditCardErrorMessage;
    _cashbackErrorMessage;
    _cardHolderName;
    _cardNumber;
    _cardType;
    _cvv;
    _expiryMonth;
    _expiryYear;
    _email;
    _telefone;
    _saldoCashback = '0.00';
    _valorCashback = 0.00;
    operacao;
    isLoading = false;
    closeCartData;

    _disablePayment = true;
    get disablePayment() {return this._disablePayment || this.isLoading;}

    @wire(MessageContext)
    messageContext;

    @wire(getRecord, {recordId: strUserId, fields: [PROFILE_NAME_FIELD]}) 
    wireuser({error, data}) {
        if(error) {
            console.log('Get User Profile ERROR => ' + error);
        }
        else if(data) {
            this.prfName = data.fields.Profile.value.fields.Name.value;        
        }
    }

    connectedCallback() {
        // this.getPublicKey();
        this.getCustomerCashback();
    }

    subscribeToMessageChannel() {
        if (!this.subscription) {
            this.subscription = subscribe(
                this.messageContext,
                checkoutMessage,
                (message) => this.handleMessage(message),
                { scope: APPLICATION_SCOPE }
            );
        }
    }

    unsubscribeToMessageChannel() {
        unsubscribe(this.subscription);
        this.subscription = null;
    }

    handleMessage(message) {
        this._disablePayment = !message.taxSuccess;
    }

    showError = false;
    @api
    checkoutMode = 1;
    @api
    error;


    @api
    checkoutSave() {
        if (!this.checkValidity) {
            throw new Error(
                            'Verificar os dados de pagamento.'
                            );
        }
    }
    

    handlePaymentType(event){
        this.value = event.detail.value;
        this._boleto = false;
        this._pix = false;
        this._credit = false;
        this._formaPagamento = this.value;
        if(this.value == 'boleto'){
            this._boleto = true;
        }else if(this.value == 'cartao de credito'){
            this._credit = true;
        }else if(this.value == 'pix'){
            this._pix = true;
        }
    }

    @api
    get boleto() {
        return this._boleto;
    }
    
    @api
    get pix() {
        return this._pix;
    }

    @api
    get creditCard() {
        return this._credit;
    }

    @api
    get cardHolderName() {
        return this._cardHolderName;
    }

    @api
    get cardType() {
        return this._cardType;
    }

    @api
    get cardNumber() {
        return this._cardNumber;
    }

    @api
    get email() {
        return this._email;
    }

    @api
    get telefone() {
        return this._telefone;
    }

    @api
    get cvv() {
        return this._cvv;
    }

    @api
    get expiryMonth() {
        return this._expiryMonth;
    }

    @api
    get expiryYear() {
        return this._expiryYear;
    }

    @api cardHolderNameRequired = false;
    @api cardTypeRequired = false;
    @api expiryMonthRequired = false;
    @api expiryYearRequired = false;
    @api cvvRequired = false;
    @api hideCardHolderName = false;
    @api hideCardType = false;
    @api hideCvv = false;
    @api hideExpiryMonth = false;
    @api hideExpiryYear = false;

    
    @api
    get saldoCashBack(){
        return this._saldoCashback;
    }

    @api
    get pixErrorMessage() {
        return this._pixErrorMessage;
    }

    set pixErrorMessage(newMessage) {
        this._pixErrorMessage = newMessage;
    }

    @api
    get cashbackErrorMessage(){
        return this._cashbackErrorMessage;
    }

    @api
    get creditCardErrorMessage() {
        return this._creditCardErrorMessage;
    }

    set creditCardErrorMessage(newMessage) {
        this._creditCardErrorMessage = newMessage;
    }

    @api
    get checkValidity() {
        return this.reportValidity();
    }

    @api
    reportValidity() {
        if(parseFloat(this._valorCashback) != 0){
            if(parseFloat(this._valorCashback) < 0 || parseFloat(this._valorCashback) > parseFloat(this._saldoCashBack)) {
                this._cashbackErrorMessage = "Valor inválio de cashback.";
                return false;
            }
        }
        if(this.creditCard){
            let incompleteFields = this.hasIncompleteCardPaymentFields();

            const componentsToValidate = this.template.querySelectorAll(
                '[data-validate]'
            );

            const validateFields = [...componentsToValidate].reduce(
                (result, component) => component.reportValidity() && result,
                true
            );

            return !incompleteFields && validateFields;
        }else if(this._pix){ 
            return true;
        }else if(this.boleto){
            return true;
        }else{
            return false;
        }
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
                this._creditCardErrorMessage = fieldRequiredValueMap[i].error;
                return true;
            }
        }

        return false;
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


    get expiryMonths() {
        const expiryMonths = [],
            noOfMonths = 12;
        for (let month = 1; month <= noOfMonths; month++) {
            expiryMonths.push({ label: month, value: month.toString() });
        }
        return expiryMonths;
    }


    get cardTypes() {
        return Object.entries(cardTypesObj).map((keyValue) => ({
            label: keyValue[1],
            value: keyValue[0]
        }));
    }


    get cardNumberClass() {
        const sldsColumnSize = this.hideCvv
            ? 'slds-size_1-of-1'
            : 'slds-size_2-of-3';
        return 'slds-form-element ' + sldsColumnSize;
    }

    handlePixEmail(event){
        this._email = event.target.value;
        event.target.reportValidity();
    }

    handlePixTelefone(event){
        this._telefone = event.target.value;
        event.target.reportValidity();
    }

    handleCardHolderNameChange(event) {
        this._cardHolderName = event.target.value;
        event.target.reportValidity();
    }

    handleCardTypeChange(event) {
        this._cardType = event.target.value;
        event.target.reportValidity();
    }

    handleCardNumberChange(event) {
        this._cardNumber = event.target.value;
        event.target.reportValidity();
    }

    handleCvvChange(event) {
        this._cvv = event.target.value;
        event.target.reportValidity();
    }

    handleExpiryMonthChange(event) {
        this._expiryMonth = event.target.value;
        event.target.reportValidity();
    }

    handleExpiryYearChange(event) {
        this._expiryYear = event.target.value;
        event.target.reportValidity();
    }

    preventSensitiveInformationPropagation(keyboardEvent) {
        keyboardEvent.stopPropagation();
    }

    orderId;
    orderNumber;

    async placeOrder() {
        if( this.reportValidity()) {
            await this.createOrder();

            let parameter = {
                orderId:        this.orderId,
                formaPagamento: this._formaPagamento,
                valorCashback:  this._valorCashback,
                operacao:       this.operacao      
            }

            if(parameter.formaPagamento.toLowerCase() === 'cartao de credito' && !this.showAlert) {
                let cardParameter = {
                    parcelas:       this._parcelas,
                    cvvCartao:      this._cvv,
                    anoCartao:      this._expiryYear,
                    mesCartao:      this.expiryMonth,
                    numeroCartao:   this.cardNumber,
                    nomeCartao:     this.cardHolderName  
                };
                let encryptedCard;
                encryptedCard = await this.encryptCard(cardParameter);
                parameter.parcelas  = cardParameter.parcelas;
                parameter.encryptedCard = encryptedCard;
            }
            // pagamento sincrono
            if(!this.showAlert){
                switch (parameter.formaPagamento.toLowerCase()) {
                    case 'boleto':
                        await this.sendOrder(parameter);
                        this.closeCartData = parameter;
                        break;
                        
                    case 'pix':
                        await this.payOrder(parameter);
                        this.operacao = 2 // payment created, not done
                        if(!this.showAlert) {
                            await this.sendOrder(parameter);
                        }
                    break;
                        
                    default:
                        this.operacao = 2 // create order in agrosys
                        await this.sendOrder(parameter);

                        if(!this.showAlert) { // pay order
                            await this.payOrder(parameter);
                        }
                        if(!this.showAlert) { // reserve order in agrosys (now operacao == 3)
                            await this.sendOrder(parameter);
                        }
                        break;
                }
                if(!this.showAlert) {
                    await this.closeWebCart(this.orderNumber, JSON.stringify(this.closeCartData)); 
                }
            }
        }
    }

    async encryptCard(cardData) {
        this.isLoading = true;
        try {
            const result = JSON.parse(await getPagBankPublicKey());
            console.log('result encryptCard OBJ => ' + result);
            console.log('result encryptCard STR => ' + JSON.stringify(result));
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

    async updatePublicKey(cardData) {
        this.isLoading = true;
        await changePagBankPublicKey().then(result => {
            result = JSON.parse(result);
            if(result.hasError) {
                this.alertMessage = result.errorMessage;
                this.showAlert = result.errorMessage;
            }
            else {
                this.encryptCard(cardData);
            }
        })
        .catch(error => {
            this.alertMessage   = error;
            this.showAlert      = true;
        });
        this.isLoading = false; 
    }

    async getCustomerCashback() {
        await getCashback().then(result => {
            let retorno = JSON.parse(result);
            if(retorno.error){
                this.alertLogId     = retorno.logId;
                this.alertMessage   = retorno.errorMessage;
                this.showAlert      = true;
                this.handleShowToast('Erro.', retorno.errorMessage, 'error');
                this.isLoading      = false;
            }
            else {
                if(retorno.cashbackBalance != null && retorno.cashbackBalance != undefined) {
                    let cashbakBalance = parseFloat(retorno.cashbackBalance.replace(",", "."));
                    if(cashbakBalance != 0){
                        this._saldoCashback = retorno.cashbackBalance;
                    }
                }
            }
        })
        .catch(error => {
            this.alertMessage   = error;
            this.showAlert      = true;
            this.isLoading      = false;
        });
        this.subscribeToMessageChannel();
    }

    async createOrder() {
        this.isLoading = true;
        await createOrder({valorCashback: this._valorCashback})
        .then(result => {
            let retorno = JSON.parse(result);
            if(retorno.error){
                this.alertLogId     = retorno.logId;
                this.alertMessage   = retorno.errorMessage;
                this.showAlert      = true;
                this.handleShowToast('Erro.', retorno.errorMessage, 'error');
                this.isLoading      = false;
            }
            this.orderId        = retorno.orderId;
            this.orderNumber    = retorno.orderNumber;
            this.operacao       = 2;
        })
        .catch(error => {
            this.alertMessage   = JSON.stringify(error);
            this.showAlert      = true;
            this.isLoading      = false;
            this.handleShowToast('Erro no processo do Checkout.', '', 'error');
        });  
    }

    async sendOrder(parameter) {
        this.isLoading = true;
        parameter.operacao = this.operacao;
        await setPedido({jsonString: JSON.stringify(parameter)})
        .then(result => {
            this.isLoading = false;
            let retorno = JSON.parse(result)
            if(retorno.hasError) {
                this.alertLogId     = retorno.logId;
                this.alertMessage   = retorno.errorMessage;
                this.showAlert      = true;
                this.handleShowToast('Erro.', retorno.errorMessage, 'error');
                this.isLoading      = false;
            }
        })
        .catch(error => {
            this.alertMessage   = JSON.stringify(error);
            this.showAlert      = true;
            this.isLoading      = false;
        })
    }

    async payOrder(parameter) {
        this.isLoading = true;
        await pay({jsonParameter: JSON.stringify(parameter)})
        .then(result => {
            let retorno = JSON.parse(result);
            if(retorno.error) {
                this.alertLogId     = retorno.logId;
                this.alertMessage   = retorno.errorMessage;
                this.showAlert      = true;
                this.isLoading      = false;
                // this.handleShowToast('Erro nos dados de pagamento.', 'Por favor verifique os dados digitados.', 'error');
            }
            this.operacao   = 3;
            this.closeCartData = retorno.paymentReturn;
            this.isLoading  = false;
        })
        .catch(error => {
            this.alertMessage   = JSON.stringify(error);
            this.showAlert      = true;
            this.isLoading      = false;
        });
    }

    async closeWebCart(orderNumber, data) {
        closeWebCart({data: data})
        .then(result => {
            
            let retorno = JSON.parse(result);
            if(retorno.hasError) {
                this.alertLogId     = retorno.logId;
                this.alertMessage   = retorno.errorMessage;
                this.showAlert      = true;
                this.isLoading      = false;
            }
            else {
                window.location.href = basePath + "/order?orderNumber=" + orderNumber + retorno.link;
                this.isLoading = false;
            }
            
        })
        .catch(error => {
            this.alertMessage   = JSON.stringify(error);
            this.showAlert      = true;
            this.isLoading      = false;
        });
    }

    handleShowToast(title, message, type) {
        this.dispatchEvent(
            new ShowToastEvent({
                title: title,
                message: message,
                variant: type
            })
            );
        }

    handleParcelas(event){
        this._parcelas = event.target.value;
    }
    
    handleCashback(event) {
        // Inicializa _valorCashback se for undefined
        if(this._valorCashback == undefined) {
            this._valorCashback = 0.00;
        }
    
        // Substitui a vírgula por ponto para garantir que o número seja válido
        let valorCashback = parseFloat(event.target.value.replace(',', '.'));
        let saldo = parseFloat(this._saldoCashback.replace(',', '.'));
    
        // Verifica se o valor inserido é menor ou igual a 0
        if(valorCashback < 0 || isNaN(valorCashback)) {
            event.target.value = 0.00;
            this._valorCashback = 0.00;
        }
        else if (valorCashback > saldo) {
            this._valorCashback = saldo;
            event.target.value = saldo.toFixed(2);
        }
        else {
            if(this._valorCashback != valorCashback) {
                this._valorCashback = valorCashback;
                event.target.value = valorCashback.toFixed(2);
            }
        }
    }
    
    @api
    get valorCashback(){
        return this._valorCashback;
    }
    
    closeAlert() {
        this.isLoading = true;
        this.showAlertMesage = '';
        this.showAlert = false;
        this.isLoading = false;
        if(this.prfName.toLowerCase() !== 'system administrator' && this.prfName.toLowerCase() !== 'administrador do sistema') {
            window.location.href = basePath + '/cart';
        }
    }
    
}