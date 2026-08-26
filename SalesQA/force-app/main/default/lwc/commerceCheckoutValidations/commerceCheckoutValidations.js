import { api, wire } from 'lwc';
import { CheckoutComponentBase } from 'commerce/checkoutApi';
import Toast from 'lightning/toast';
import ToastContainer from 'lightning/toastContainer';
import { CartSummaryAdapter } from 'commerce/cartApi';
import { isInSitePreview } from "c/utils";
import getCheckoutAlerts from '@salesforce/apex/CommerceCheckoutValidationController.getCheckoutAlerts';
import updateCart from '@salesforce/apex/CommerceCustomCartItemsController.updateCart';

const CheckoutStage = {
    CHECK_VALIDITY_UPDATE: 'CHECK_VALIDITY_UPDATE',
    REPORT_VALIDITY_SAVE: 'REPORT_VALIDITY_SAVE',
    BEFORE_PAYMENT: 'BEFORE_PAYMENT',
    PAYMENT: 'PAYMENT',
    BEFORE_PLACE_ORDER: 'BEFORE_PLACE_ORDER',
    PLACE_ORDER: 'PLACE_ORDER'
};

export default class CommerceCheckoutValidations extends CheckoutComponentBase {

    @api errorMessage;

    _userCheckoutAllowed = false;
    isPreview = false;
    isLoading = true;
    savePending = false;
    condicoes;
    condicaoSelecionada;
    dataEntrega;
    obsNF = '';
    numeroPedidoCliente = '';
    cartId;
    dataMinimaEntrega;
    dataMaximaEntrega
    janelaCarregamento = '';
    permiteRcs = false;
    rcs;
    permiteRemessaDeposito = false;
    existePedido = false;
    checkoutAlerts;
 
    connectedCallback() {
        this.isPreview = isInSitePreview();
        const toastContainer = ToastContainer.instance();
        toastContainer.maxToasts = 5;
        toastContainer.toastPosition = 'top-center';
    }

    stageAction(checkoutStage) {
        switch (checkoutStage) {
            case CheckoutStage.BEFORE_PLACE_ORDER:
                return Promise.resolve(this.reportValidity());
            default:
                return Promise.resolve(true);
        }
    }

    get disabledRcsCheckbox(){
        return true;
    }

    @wire(CartSummaryAdapter, {})
    cartSummary({ error, data }) {
        if (data && data.customFields) {
            const currentDate = new Date();
            this.cartId = data.cartId;
            this.dataMinimaEntrega = data.customFields[0]?.DataMinimaEntrega__c ?? currentDate;
            this.dataMaximaEntrega = data.customFields[0]?.DataMaximaEntrega__c ?? currentDate;
            this.janelaCarregamento = data.customFields[0]?.JanelaCarregamento__c ?? '';
            this.dataEntrega = data.customFields[0]?.DataEntrega__c;
            this.permiteRcs = data.customFields[0]?.PermiteVendaContaOrdem__c; // Campo que define se a Conta permite RCS
            this.rcs = data.customFields[0]?.VendaContaOrdem__c; // Campo que define se este pedido deve ser RCS
            this.permiteRemessaDeposito = data.customFields[0]?.PermiteRemessaDeposito__c;
            this.obsNF = data.customFields[0]?.ObservacaoNotaFiscal__c;
            this.numeroPedidoCliente = data.customFields[0]?.NumeroPedidoCliente__c;
            this.existePedido = data?.customFields[0]?.PedidoMesmoDiaMesmoCliente__c;

            getCheckoutAlerts({
                cartId: this.cartId
            }).then(result => {
                this.checkoutAlerts = result;
                this._userCheckoutAllowed = result.isCheckoutValid;
                this.condicoes = result.condicoesPagamento;
                this.condicaoSelecionada = data.customFields[0]?.CodigoCondicaoPagamentoSelecionada__c;
            }).catch(error => {
                console.error('CommerceCheckoutValidations - getCheckoutAlerts: ', error);
            }).finally(f =>{
                this.isLoading = false;
            })

        } else if (error) {
            console.error(error);
        }
    }

    async reportValidity() {
        this.dispatchUpdateErrorAsync({
            groupId: 'Security'
        });

        if(!this.refs.inputDataEntrega.reportValidity()){
            Toast.show({
                label: 'Informe uma Data de Carregamento válida',
                variant: 'warning',
                mode: 'dismissible'
            }, this);
        } else {
            if(this.existePedido && !this.refs.inputExistePedido.checked){
                Toast.show({
                    label: 'Cliente já possui pedido, confirme que deseja continuar',
                    variant: 'warning',
                    mode: 'dismissible'
                }, this);
            } else {
                if(this.savePending){
                    Toast.show({
                        label: 'Aguarde até que todos os dados preenchidos sejam gravados antes de gerar o pedido',
                        variant: 'info',
                        mode: 'dismissible'
                    }, this);
                } else {
                    if (!this._userCheckoutAllowed) {
                        Toast.show({
                            label: 'Pedido não realizado, por gentileza verfique as inconsistências e {0}',
                            labelLinks: [{
                                url: '/forcadevendasavivar/cart',
                                label: "atualize o carrinho",
                            }],
                            variant: 'warning',
                            mode: 'dismissible'
                        }, this);
                    }
                }
            }
        }

        return (this._userCheckoutAllowed && 
                !this.savePending &&
                this.refs.inputDataEntrega.checkValidity() &&
                (this.existePedido == false || this.refs.inputExistePedido.checked) );
    }

    handleChangeDataEntrega(event){
        event.target.setCustomValidity('');
        if(this.refs.inputDataEntrega.reportValidity()){
            if(this.janelaCarregamento !== ''){
                let parts = event.target.value.split('-');
                let date = new Date(parts[0], parts[1] - 1, parts[2]);
                let gmtToday = new Date();
                gmtToday.setHours(0, 0, 0, 0);
                const dayOfWeek = date.getDay();

                if(this.janelaCarregamento.includes(dayOfWeek)) {
                    event.target.setCustomValidity('');
                } else {
                    event.target.setCustomValidity('Data informada está fora da janela de carregamento');
                }
            }

            if(event.target.reportValidity()){
                this.doUpdate();
            }
        }
    }

    handleCondicaoPagamentoChange() {
        if(this.refs.inputCondicaoPagamento.reportValidity()){
            this.doUpdate();
        }
    }

    handleChangeObsNF(event){
        const input = event.target;
        const sanitized = this.sanitizeInput(input.value);
        input.value = sanitized;

        this.savePending = true;
        clearTimeout(this.tempo);
        this.tempo = setTimeout(() => {
            if(this.refs.inputDataEntrega.reportValidity()){
                this.doUpdate();
            }
        }, 2000);
    }

    handleChangeNumeroPedido(event){
        const input = event.target;
        const sanitized = this.sanitizeInputPedido(input.value);
        input.value = sanitized;

        this.savePending = true;
        clearTimeout(this.tempo);
        this.tempo = setTimeout(() => {
            if(this.refs.inputDataEntrega.reportValidity()){
                this.doUpdate();
            }
        }, 2000);
    }

    handleChangeExistePedido(event){
        this.savePending = true;
        clearTimeout(this.tempo);
        this.tempo = setTimeout(() => {
            if(this.refs.inputDataEntrega.reportValidity()){
                this.doUpdate();
            }
        }, 200);
    }

    handleChangeRCS(event){
        console.log('handleChangeRCS', event);

        this.savePending = true;
        clearTimeout(this.tempo);
        this.tempo = setTimeout(() => {
            if(this.refs.inputDataEntrega.reportValidity()){
                this.doUpdate();
            }
        }, 1000);
    }

    async doUpdate(){
        clearTimeout(this.globalTime);
        this.globalTime = setTimeout(() => {
            this.refs.inputDataEntrega.disabled = true;
            this.refs.inputObsNF.disabled = true;
            this.refs.inputNumeroPedidoCliente.disabled = true;
            this.refs.inputCondicaoPagamento.disabled = true;
            
            if(this.refs.inputRcs){
                this.refs.inputRcs.disabled = true;
            }

            if(this.refs.inputExistePedido){
                this.refs.inputExistePedido.disabled = true;
            }

            updateCart({
                activeCartOrId: this.cartId,
                dataEntrega: this.refs.inputDataEntrega.value,
                rcs: this.refs.inputRcs ? this.refs.inputRcs.checked : false,
                obsNF: this.refs.inputObsNF.value,
                codigoCondicaoPagamento: this.refs.inputCondicaoPagamento.value,
                numeroPedidoCliente: this.refs.inputNumeroPedidoCliente.value,
                existePedido: this.refs.inputExistePedido ? this.refs.inputExistePedido.checked : false
            }).then(result => {

                getCheckoutAlerts({
                    cartId: this.cartId
                }).then(result => {
                    this.checkoutAlerts = result;
                    this._userCheckoutAllowed = result.isCheckoutValid;
                }).catch(error => {
                    console.error('CommerceCheckoutValidations - getCheckoutAlerts: ', error);
                }).finally(f =>{
                    this.isLoading = false;
                })

            }).catch(error => {
                console.error('CommerceCheckoutValidations - updateCart: ', error);
            }).finally(f =>{
                this.isLoading = false;
                this.savePending = false;
                this.refs.inputDataEntrega.disabled = false;
                this.refs.inputObsNF.disabled = false;
                this.refs.inputNumeroPedidoCliente.disabled = false;
                this.refs.inputCondicaoPagamento.disabled = false;

                if(this.refs.inputRcs){
                    this.refs.inputRcs.disabled = false;
                }

                if(this.refs.inputExistePedido){
                    this.refs.inputExistePedido.disabled = false;
                }
            })
        }, 500);
    }

    sanitizeInput(value) {
        if (!value) return '';

        return value
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "") // remove acentos
            .toUpperCase()
            .replace(/[^A-Z0-9\/\. ]/g, ""); // mantém apenas válidos
    }

    handleKeyPress(event) {
        const char = event.key;
        if (char.length > 1) return;
        if (!/^[A-Z0-9\/\. ]$/.test(char.toUpperCase())) {
            event.preventDefault();
        }
    }

    handleInput(event) {
        const input = event.target;
        const sanitized = this.sanitizeInput(input.value);
        if (input.value !== sanitized) {
            input.value = sanitized;
        }
    }

    sanitizeInputPedido(value) {
        if(!value){
            return '';  
        } else {
            return value
                    .normalize("NFD")
                    .replace(/[\u0300-\u036f]/g, "") // remove acentos
                    .toUpperCase()
                    .replace(/[^A-Z0-9\/\.\- ]/g, ""); // mantém apenas válidos
        }
    }

    handleKeyPressPedido(event) {
        const char = event.key;
        if (char.length > 1) {
            return;
        } else {
            if (!/^[A-Z0-9\/\.\- ]$/.test(char.toUpperCase())) {
                event.preventDefault();
            }
        }
    }

    handleInputPedido(event) {
        const input = event.target;
        const sanitized = this.sanitizeInputPedido(input.value);
        if (input.value !== sanitized) {
            input.value = sanitized;
        }
    }

}