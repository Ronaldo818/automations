import { LightningElement, wire, api, track } from 'lwc';
import {getRecord} from 'lightning/uiRecordApi';
import {publish, MessageContext} from "lightning/messageService";
import PROFILE_NAME_FIELD from '@salesforce/schema/User.Profile.Name';
import strUserId from '@salesforce/user/Id';
import enderecoEntrega from '@salesforce/apex/B2BCheckoutController.enderecoEntrega';
import criarCartDeliveryGroup from '@salesforce/apex/B2BCheckoutController.criarCartDeliveryGroup';
export default class B2BCheckout extends LightningElement {
    @wire(MessageContext)
    messageContext;
    @api recordId;
    @api minimumOrderValue;
    @api messageMinimumOrderError;
    @track prfName;
    @track isLoading;
    @track error;
    @track cidade;
    @track unidade;
    @track contatoId;
    @track contaId;
    @track cartId;
    @track estoque = {};
    @track calcImp;
    @track calcImpError = false;
    @track valorInferiorPedidoMinimo = false;
    @track cid;
    @track rua;
    @track cep;
    @track estado;
    @track bairro;
    @track complemento;
    @track numero;
    @track semEstoque = false;
    @track showAlert = false;
    @track showAlertMesage;
    @track resumoCompra = false;

    @wire(getRecord, {recordId: strUserId, fields: [PROFILE_NAME_FIELD]}) 
    wireuser({error, data}) {
        if(error) {
            console.log('Get User Profile ERROR => ' + error);
        }
        else if(data) {
            this.prfName = data.fields.Profile.value.fields.Name.value;        
        }
    }

    step(steps) {
        if(!this.showAlert && steps === 1) this.getEndereco();
        if(!this.showAlert && steps === 3) this.getCriarCartDeliveryGroup();
    }

    connectedCallback() {
        this.step(1);
    }

    // Step 1
    getEndereco() {
        this.isLoading = true;
        enderecoEntrega()
        .then(result => {
            let retorno = JSON.parse(result);
            let unidade;
            if(retorno.hasError) {
                this.showAlertMesage = retorno.errorMesage;
                this.showAlert = true;
            }
            else {
                this.cid = retorno.endereco.cidade;
                this.rua = retorno.endereco.rua;
                this.cep = retorno.endereco.cep;
                this.estado = retorno.endereco.estado;
                this.bairro = retorno.endereco.bairro;
                this.complemento = retorno.endereco.complemento;
                this.numero = retorno.endereco.numero;
                unidade = retorno.endereco.unidade;
            }
            this.isLoading = false;

            if(unidade === 13) {
                this.step(2);
            }
            else {
                this.step(3);
            }
        })
        .catch(error => {
            let erro = error;
            console.log('ENDEREÇO ERROR ==>');
            console.log(erro);
            this.showAlertMesage = 'Ocorreu um erro ao consultar seu endereço de entrega. Por favor, tente novamente mais tarde. Se o problema persistir informe ao seu representante de vendas ou administrador.';
            this.showAlert = true;
            this.isLoading = false;
        });
    }

    // Step 2
    verificarEstoque() {
        this.isLoading = true;
        verificarEstoque()
        .then(result => {
            this.estoque = JSON.parse(result);
            this.semEstoque = this.estoque.removeuItem;
            this.step(3);
            this.isLoading = false;
        })
        .catch(error => {
            this.error = JSON.stringify(error);
            console.error('Error function verificarEstoque()');
            this.isLoading = false;
        })
    }

    // Step 3
    getCriarCartDeliveryGroup(){
        this.isLoading = true;
        criarCartDeliveryGroup().then(result => {
            // this.step(4);
            this.isLoading = false;
        })
        .catch(error => {
            this.error = JSON.stringify(error);
            console.error('Error function getCriarCartDeliveryGroup()');
            this.isLoading = false;
        });
    }

    closeAlert() {
        this.isLoading = true;
        this.showAlertMesage = '';
        this.showAlert = false;
        this.isLoading = false;
        if(this.prfName.toLowerCase() !== 'system administrator' && this.prfName.toLowerCase() !== 'administrador do sistema') {
            window.location.href = window.location.origin + '/cart';
        }
    }

    closeAlertStock() {
        this.semEstoque = false;    
    }
}