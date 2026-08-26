import { LightningElement, api, track, wire } from 'lwc';
import { getObjectInfo } from 'lightning/uiObjectInfoApi';
import { getPicklistValues } from 'lightning/uiObjectInfoApi';
import { validateCpf, validateCnpj, validateEmail, validateCep, maskCpf, maskPhone, maskCep, maskCnpj } from "c/utils";
import Toast from 'lightning/toast';
import ToastContainer from 'lightning/toastContainer';
import getAddressByCEP from '@salesforce/apex/B2BCommerceRegistrationController.getAddressByCEP';
import veryifyRecord from '@salesforce/apex/B2BCommerceRegistrationController.veryifyRecord';
import saveRecord from '@salesforce/apex/B2BCommerceRegistrationController.saveRecord';

import AccountExist from '@salesforce/apex/CommerceUtil.AccountExist';

import ACCOUNT_OBJECT from '@salesforce/schema/Account';
import ESTADO_FIELD from '@salesforce/schema/Account.EstadoComercial__c';
import CATEGORIA_FIELD from '@salesforce/schema/Account.Categoria__c';
import SUBCATEGORIA_FIELD from '@salesforce/schema/Account.Subcategoria__c';
import RAMO_ATIVIDADE_FIELD from '@salesforce/schema/Account.RamoAtividade__c';

export default class B2BCommerceRegistration extends LightningElement {

    @api initialText;
    @track isLoading = false;

    showSpinner = false;
    step = 1;
    totalSetps = 0;
    tipoCadastro = '';
    identificacaoFiscal = '';
    showSpinner = false;
    mapFiels = new Map();

    connectedCallback() {
        const toastContainer = ToastContainer.instance();
        toastContainer.maxToasts = 5;
        toastContainer.toastPosition = 'top-center';
    }

    @wire(getObjectInfo, { objectApiName: ACCOUNT_OBJECT })
    accountObjectInfo;

    @wire(getPicklistValues, { recordTypeId: '$accountObjectInfo.data.defaultRecordTypeId', fieldApiName: ESTADO_FIELD })
    estadoOptions;

    @wire(getPicklistValues, { recordTypeId: '$accountObjectInfo.data.defaultRecordTypeId', fieldApiName: CATEGORIA_FIELD })
    categoriaOptions;

    @wire(getPicklistValues, { recordTypeId: '$accountObjectInfo.data.defaultRecordTypeId', fieldApiName: SUBCATEGORIA_FIELD })
    subcategoriaOptions;

    @wire(getPicklistValues, { recordTypeId: '$accountObjectInfo.data.defaultRecordTypeId', fieldApiName: RAMO_ATIVIDADE_FIELD })
    ramoAtividadeOptions;

    get CidadeEntregaFieldName(){
        return this.accountObjectInfo.data.fields.CidadeEntrega__c.label;
    }

    get currentStep(){
        return this.step; // + ' / ' + mapStepPages.size;
    }

    get isPJ (){
        return (this.tipoCadastro == 'pj');
    }

    get showStep1(){
        return this.step === 1;
    }

    get showStep2(){
        return this.step === 2;
    }

    get showStep3(){
        return this.step === 3;
    }

    get showStep4(){
        return this.step === 4;
    }

    get showFooter() {
        return (this.step !== 1 && this.step !== 4);
    }

    get labelNext(){
        if(this.step == 3){
            return 'Confirmar';
        }
        else{
            return 'Próximo';
        }
    }

    checkFormValidity(){
        return [...this.template.querySelectorAll('lightning-input, lightning-combobox, lightning-input-field')]
        .reduce((validSoFar, inputFieldReference) => {
            inputFieldReference.reportValidity();
            return validSoFar && inputFieldReference.reportValidity();
        }, true);
    }

    handleGenericChange(event){
        // console.log('subcategoriaOptions', this.subcategoriaOptions);
    }

    handleClickTipoCadastro(event){
        console.log('Tipo cadastro', event.currentTarget.value);
        this.tipoCadastro = event.currentTarget.value;
        this.step++;
    }

    handleClickVoltarInicio(Event){
        this.identificacaoFiscal = '';
        this.mapFiels.clear();
        this.step = 1;
        this.showSpinner = false;
    }

    handleClickPreviousStep(event){
        if(this.step > 1){
            this.step--;
        }
    }

    isIdentificationInputDisabled = false;

    async handleClickNextStep(event){

        if(this.step == 2){
            if(this.checkFormValidity()){
                if(this.refs.inputCPF){
                    this.identificacaoFiscal = this.refs.inputCPF.value;
                } else if(this.refs.inputCNPJ) {
                    this.identificacaoFiscal = this.refs.inputCNPJ.value;
                }
            } else {
                this.showToast('Informe uma identificação fiscal válida','', 'warning');
                return;
            }
            this.isIdentificationInputDisabled = true;
            if(await AccountExist({identification: this.identificacaoFiscal})){
                this.showToast('Conta já cadastrada no sistema','', 'warning');
                this.isIdentificationInputDisabled = false;
                return;
            }
            else{
                this.isIdentificationInputDisabled = false;
            }

        } else if(this.step == 3){
            if(!this.checkFormValidity()){
                this.showToast('Preencha todos os campos obrigatórios de forma adequada','', 'warning');
                return;
            } else {

                this.showSpinner = true;
                const formInputs = this.template.querySelectorAll('lightning-input, lightning-combobox, lightning-input-field, lightning-record-picker');
                formInputs.forEach(element => {
                    // console.log('element.fieldApiName | element.value', element.fieldApiName + ' | ' + element.value );
                    this.mapFiels.set(element.fieldApiName, element.value);
                });

                console.log('sending:', Object.fromEntries(this.mapFiels));

                await saveRecord({
                    fields: JSON.stringify(Object.fromEntries(this.mapFiels))
                })
                .then(result => {
                    console.log('Create Account', result);
                    this.showSpinner = false;

                    window.scrollTo({
                        top: 0,
                        left: 0,    
                        behavior: 'smooth'
                    });

                    this.step++;

                })
                .catch(error => {
                    this.showToastError('Ocorreu um erro ao criar a conta', error);
                    return;
                })

            }
        }

        if(this.step < 3){
            this.step++;
        }

        if(this.step == 3){
            console.log('identificacaoFiscal', this.identificacaoFiscal);
            clearTimeout(this.tempo);
            this.tempo = setTimeout(() => {
                this.refs.inputName.focus();

                if(this.refs.inputCNPJ || this.refs.inputCPF) {
                    this.showSpinner = true;
                    try {
                        veryifyRecord({
                            identificacaoFiscal: this.identificacaoFiscal
                        })
                        .then(result => {
                            console.log('Result', result);

                            if(result){
                                if(result.recordData){
                                    console.log('Achou');
                                    this.refs.inputName.value = result.recordData.Name;
                                    this.refs.inputFantasia.value = result.recordData.NomeFantasia__c;
                                    this.refs.inputCep.value = result.recordData.CEPEntrega__c;
                                    this.refs.inputRua.value = result.recordData.RuaEntrega__c;
                                    this.refs.inputNumero.value = result.recordData.NumeroEntrega__c;
                                    this.refs.inputBairro.value = result.recordData.BairroEntrega__c;
                                    this.refs.inputComplemento.value = result.recordData.Complemento__c;
                                    this.refs.inputMunicipio.value = result.recordData.CidadeEntrega__c;
                                    this.refs.inputEstado.value = result.recordData.EstadoEntrega__c;

                                    this.refs.inputEmail.value = result.recordData.Email__c;
                                    this.refs.inputFone.value = result.recordData.TelefoneComercial__c;

                                    this.refs.inputAnoFundacao.value = result.recordData.AnoFundacao__c;
                                    //this.refs.inputRamoAtividade.value = result.recordData.RamoAtividade__c;
                                    this.refs.inputCategoria.value = result.recordData.Categoria__c;
                                    this.refs.inputSubcategoria.value = result.recordData.Subcategoria__c;
                                } else {
                                    if(result.cnpjData){
                                        this.refs.inputName.value = result.cnpjData.nome;
                                        this.refs.inputFantasia.value = result.cnpjData.fantasia;
                                        this.refs.inputAnoFundacao.value = result.cnpjData.c_ano_fundacao;
                                        this.refs.inputCep.value = (result.cnpjData.cep ? maskCep(result.cnpjData.cep) : '');
                                        this.refs.inputRua.value = result.cnpjData.logradouro;
                                        this.refs.inputNumero.value = result.cnpjData.numero;
                                        this.refs.inputBairro.value = result.cnpjData.bairro;
                                        this.refs.inputComplemento.value = result.cnpjData.complemento;
                                        this.refs.inputMunicipio.value = result.cnpjData.municipio;
                                        this.refs.inputEstado.value = result.cnpjData.uf
                                        
                                        console.log('result.cnpjData.email',result.cnpjData.email);
                                        this.refs.inputEmail.value = result.cnpjData.email;

                                        console.log('result.cnpjData.telefone',result.cnpjData.telefone);
                                        this.refs.inputFone.value = result.cnpjData.telefone;
                                    }
                                }
                            }
                            this.showSpinner = false;
                        })
                    } catch (ex) {
                        this.showToastError('Ocorreu um erro inesperado', ex);
                    }
                }
            }, 10);
        }
    }

    handleChangeCpf(event){
        event.target.value = maskCpf(event.target.value);
        if(validateCpf(event.target.value)){
            event.target.setCustomValidity('');
        } else {
            event.target.setCustomValidity('CPF inválido');
        }
        event.target.reportValidity();
    }

    handleChangeCnpj(event){
        event.target.value = maskCnpj(event.target.value);
        if(validateCnpj(event.target.value)){
            event.target.setCustomValidity('');
        } else {
            event.target.setCustomValidity('CNPJ inválido');
        }
        event.target.reportValidity();
    }

    handleChangeEmail(event){
        if(validateEmail(event.target.value)){
            event.target.setCustomValidity('');
        } else {
            event.target.setCustomValidity('Endereço de email inválido');
        }
        event.target.reportValidity();
    }

    handleChangeTelefone(event){
        event.target.value = maskPhone(event.target.value);
        event.target.reportValidity();
    }

    handleChangeCep(event){
        console.log('handleChangeCep');
        var cep = event.target.value;
        event.target.value = maskCep(cep);

        clearTimeout(this.tempo);
        this.tempo = setTimeout(() => {
            var fieldApiName = event.target.fieldApiName;
            this.mapFiels.set(fieldApiName, cep);

            if(validateCep(cep)){
                this.isLoading = true;
                getAddressByCEP({ cep: cep })
                .then(result => {
                    console.log('CEP result', result);
                    if(result){
                        if(result.erro){
                            this.showToast('O CEP informado não foi encontrado, por favor preencha o endereço manualmente', '', 'warning')
                        } else {
                            this.refs.inputRua.value = result.logradouro;
                            this.refs.inputBairro.value = result.bairro;
                            this.refs.inputComplemento.value = result.complemento;
                            this.refs.inputMunicipio.value = result.localidade;
                            this.refs.inputEstado.value = result.estado;
                            this.refs.inputNumero.focus();
                        }
                    }
                    this.isLoading = false;
                })
                .catch(error => {
                    this.isLoading = false;
                    console.error(error);
                });
            }
        }, 500);
    }

    showToast(label, message, variant){
        this.showSpinner = false;
        Toast.show({label: label, message: message, variant: variant, mode: 'dismissible'}, this);
    }

    showToastError(label, error) {
        console.error(error);
        this.showSpinner = false;
        var message = '';

    if (!Array.isArray(error)) {
        error = [error];
    }
 
    message = 
        error
            // Remove null/undefined items
            .filter((error) => !!error)
            // Extract an error message
            .map((error) => {
                // UI API read errors
                if (error.body.duplicateResults && error.body.duplicateResults.length > 0) {
                    return error.body.duplicateResults.map((e) => e.message);
                }

                else if (error.body.fieldErrors && error.body.fieldErrors.length > 0 && Array.isArray(error.body.fieldErrors)) {
                    return error.body.fieldErrors.map((e) => e.message);
                }

                else if (error.body.pageErrors && error.body.pageErrors.length > 0 && Array.isArray(error.body.pageErrors)) {
                    return error.body.pageErrors.map((e) => e.message);
                }

                else if (Array.isArray(error.body)) {
                    return error.body.map((e) => e.message);
                }
                // UI API DML, Apex and network errors
                else if (error.body && typeof error.body.message === 'string') {
                    return error.body.message;
                }
                // JS errors
                else if (typeof error.message === 'string') {
                    return error.message;
                }
                // Unknown error shape so try HTTP status text
                return error.statusText;
            })
            // Flatten
            .reduce((prev, curr) => prev.concat(curr), [])
            // Remove empty strings
            .filter((message) => !!message);
    
        console.log('message', message);

        // if(error.message) {
        //     message = error.message;
        // } else if (error.body.message) {
        //     message = error.body.message;
        // }

        Toast.show({label: label, message: message[0], variant: 'error', mode: 'dismissible'}, this);
    }

}