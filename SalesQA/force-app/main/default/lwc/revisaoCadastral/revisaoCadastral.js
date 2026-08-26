import { LightningElement, track, wire } from 'lwc';
import { validateCnpj, validateCpf, validateEmail, maskPhone, maskCnpj, maskCpf } from 'c/utils';
import Toast from 'lightning/toast';
import ToastContainer from 'lightning/toastContainer';
import saveRevisaoCadastral from '@salesforce/apex/RevisaoCadastralController.saveRevisaoCadastral';
import saveFiles from '@salesforce/apex/RevisaoCadastralController.saveFiles';
import findAccountByCnpj from '@salesforce/apex/RevisaoCadastralController.findAccountByCnpj';
import findAccountByCpf from '@salesforce/apex/RevisaoCadastralController.findAccountByCpf';
import getTipoRevisaoOptions from '@salesforce/apex/RevisaoCadastralController.getTipoRevisaoOptions';

// const TIPO_REVISAO_OPTIONS = [
//     { label: 'Razão social',                       value: 'Razão social' },
//     { label: 'CNPJ',                               value: 'CNPJ' },
//     { label: 'IE/IM',                              value: 'IE/IM' },
//     { label: 'Endereço',                           value: 'Endereço' },
//     { label: 'Segmento',                           value: 'Segmento' },
//     { label: 'Contribuinte de ICMS',               value: 'Contribuinte de ICMS' },
//     { label: 'Calcula ST (industrializar ou revender)', value: 'Calcula ST (industrializar ou revender)' },
//     { label: 'Regime Especial',                    value: 'Regime Especial' }
// ];

export default class RevisaoCadastral extends LightningElement {

    @track showSpinner = false;
    @track showSuccess = false;
    @track uploadedFiles = [];
    @track isSearchingConta = false;
    @track tipoPessoa = 'PJ';

    selectedTipoRevisao = new Set();
    contaId = null;
    @track contaNome = null;
    @track contaNotFound = false;
    _contaSearchTimeout;

    @track tipoRevisaoOptions = [];

    get isPJ() {
        return this.tipoPessoa === 'PJ';
    }

    get isPF() {
        return this.tipoPessoa === 'PF';
    }

    @wire(getTipoRevisaoOptions)
    wiredTipoRevisaoOptions({ error, data }) {
        if (data) {
            this.tipoRevisaoOptions = data;
        } else if (error) {
            console.error('Erro ao carregar opções de Tipo de Revisão:', error);
        }
    }

    connectedCallback() {
        const toastContainer = ToastContainer.instance();
        toastContainer.maxToasts = 5;
        toastContainer.toastPosition = 'top-center';
    }

    get hasFiles() {
        return this.uploadedFiles.length > 0;
    }

    handleTipoRevisaoChange(event) {
        const value = event.target.value;
        if (event.target.checked) {
            this.selectedTipoRevisao.add(value);
        } else {
            this.selectedTipoRevisao.delete(value);
        }
    }

    handleTipoPessoaChange(event) {
        const tipo = event.target.dataset.tipo;
        if (!event.target.checked) {
            event.target.checked = true;
            return;
        }
        if (tipo === this.tipoPessoa) {
            return;
        }
        this.tipoPessoa = tipo;
        this.contaId = null;
        this.contaNome = null;
        this.contaNotFound = false;
        this.isSearchingConta = false;
        clearTimeout(this._contaSearchTimeout);

        const inputCnpj = this.template.querySelector('[data-id="inputCNPJ"]');
        const inputCpf = this.template.querySelector('[data-id="inputCPF"]');
        if (inputCnpj) {
            inputCnpj.value = null;
            inputCnpj.setCustomValidity('');
            inputCnpj.reportValidity();
        }
        if (inputCpf) {
            inputCpf.value = null;
            inputCpf.setCustomValidity('');
            inputCpf.reportValidity();
        }
    }

    handleChangeCnpj(event) {
        event.target.value = maskCnpj(event.target.value);
        if (event.target.value && !validateCnpj(event.target.value)) {
            event.target.setCustomValidity('CNPJ inválido');
            this.contaNome = null;
            this.contaNotFound = false;
            this.contaId = null;
            this.isSearchingConta = false;
        } else {
            event.target.setCustomValidity('');
            this.searchAccountByCnpj(event.target.value);
        }
        event.target.reportValidity();
    }

    handleChangeCpf(event) {
        event.target.value = maskCpf(event.target.value);
        if (event.target.value && !validateCpf(event.target.value)) {
            event.target.setCustomValidity('CPF inválido');
            this.contaNome = null;
            this.contaNotFound = false;
            this.contaId = null;
            this.isSearchingConta = false;
        } else {
            event.target.setCustomValidity('');
            this.searchAccountByCpf(event.target.value);
        }
        event.target.reportValidity();
    }

    searchAccountByCnpj(cnpjValue) {
        clearTimeout(this._contaSearchTimeout);
        const cnpjClean = cnpjValue ? cnpjValue.replace(/[^A-Z0-9a-z]/g, '').toUpperCase() : '';

        this.contaNome = null;
        this.contaNotFound = false;
        this.contaId = null;

        if (cnpjClean.length < 14) {
            this.isSearchingConta = false;
            return;
        }

        this.isSearchingConta = true;
        this._contaSearchTimeout = setTimeout(async () => {
            try {
                const result = await findAccountByCnpj({ cnpj: cnpjClean });
                this.applyContaResult(result);
            } catch (error) {
                this.applyContaResult(null);
            }
        }, 500);
    }

    searchAccountByCpf(cpfValue) {
        clearTimeout(this._contaSearchTimeout);
        const cpfClean = cpfValue ? cpfValue.replace(/\D/g, '') : '';

        this.contaNome = null;
        this.contaNotFound = false;
        this.contaId = null;

        if (cpfClean.length < 11) {
            this.isSearchingConta = false;
            return;
        }

        this.isSearchingConta = true;
        this._contaSearchTimeout = setTimeout(async () => {
            try {
                const result = await findAccountByCpf({ cpf: cpfClean });
                this.applyContaResult(result);
            } catch (error) {
                this.applyContaResult(null);
            }
        }, 500);
    }

    applyContaResult(result) {
        this.isSearchingConta = false;
        if (result) {
            this.contaId = result.Id;
            this.contaNome = result.Name;
            this.contaNotFound = false;
        } else {
            this.contaId = null;
            this.contaNome = null;
            this.contaNotFound = true;
        }
    }

    handleChangeTelefone(event) {
        event.target.value = maskPhone(event.target.value);
        event.target.reportValidity();
    }

    handleChangeEmail(event) {
        if (event.target.value && !validateEmail(event.target.value)) {
            event.target.setCustomValidity('Endereço de email inválido');
        } else {
            event.target.setCustomValidity('');
        }
        event.target.reportValidity();
    }

    handleFileChange(event) {
        const files = event.target.files;
        if (files && files.length > 0) {
            const newFiles = [...this.uploadedFiles];
            for (let i = 0; i < files.length; i++) {
                if (!newFiles.some(f => f.name === files[i].name)) {
                    newFiles.push(files[i]);
                }
            }
            this.uploadedFiles = newFiles;
        }
    }

    handleRemoveFile(event) {
        const fileName = event.currentTarget.dataset.name;
        this.uploadedFiles = this.uploadedFiles.filter(f => f.name !== fileName);
    }

    handleReset() {
        this.showSuccess = false;
        this.showSpinner = false;
        this.uploadedFiles = [];
        this.selectedTipoRevisao.clear();
        this.contaId = null;
        this.contaNome = null;
        this.contaNotFound = false;
        this.isSearchingConta = false;
        this.tipoPessoa = 'PJ';

        const inputs = this.template.querySelectorAll('lightning-input, lightning-textarea');
        inputs.forEach(input => {
            input.value = null;
        });

        const checkboxes = this.template.querySelectorAll('input[type="checkbox"]');
        checkboxes.forEach(cb => { cb.checked = false; });
    }

    checkFormValidity() {
        let isValid = true;

        const inputs = this.template.querySelectorAll('lightning-input, lightning-textarea');
        inputs.forEach(input => {
            input.reportValidity();
            if (!input.reportValidity()) {
                isValid = false;
            }
        });

        if (!this.contaId) {
            const msg = this.isPJ
                ? 'Informe um CNPJ válido vinculado a uma conta'
                : 'Informe um CPF válido vinculado a uma conta';
            this.showToast(msg, '', 'warning');
            isValid = false;
        }

        if (this.selectedTipoRevisao.size === 0) {
            this.showToast('Selecione ao menos um Tipo de Revisão', '', 'warning');
            isValid = false;
        }

        return isValid;
    }

    async handleSubmit() {
        if (!this.checkFormValidity()) {
            return;
        }

        this.showSpinner = true;

        try {
            const fields = {};
            const formInputs = this.template.querySelectorAll('lightning-input, lightning-textarea');
            formInputs.forEach(element => {
                if (element.fieldApiName) {
                    fields[element.fieldApiName] = element.value;
                }
            });

            fields['Conta__c'] = this.contaId;
            fields['TipoRevisao__c'] = [...this.selectedTipoRevisao].join(';');

            const recordId = await saveRevisaoCadastral({ fieldsJson: JSON.stringify(fields) });

            if (this.uploadedFiles.length > 0) {
                await this.uploadFiles(recordId);
            }

            this.showSpinner = false;
            this.showSuccess = true;

            window.scrollTo({ top: 0, left: 0, behavior: 'smooth' });

        } catch (error) {
            this.showSpinner = false;
            this.showToastError('Ocorreu um erro ao enviar a revisão cadastral', error);
        }
    }

    async uploadFiles(parentId) {
        const filePromises = this.uploadedFiles.map(file => {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = () => {
                    const base64 = reader.result.split(',')[1];
                    resolve({ fileName: file.name, base64Data: base64 });
                };
                reader.onerror = reject;
                reader.readAsDataURL(file);
            });
        });

        const fileDataList = await Promise.all(filePromises);

        for (const fileData of fileDataList) {
            await saveFiles({
                parentId: parentId,
                fileName: fileData.fileName,
                base64Data: fileData.base64Data
            });
        }
    }

    showToast(label, message, variant) {
        this.showSpinner = false;
        Toast.show({ label: label, message: message, variant: variant, mode: 'dismissible' }, this);
    }

    showToastError(label, error) {
        console.error(error);
        this.showSpinner = false;
        let message = '';
        if (error && error.message) {
            message = error.message;
        } else if (error && error.body && error.body.message) {
            message = error.body.message;
        }
        Toast.show({ label: label, message: message, variant: 'error', mode: 'dismissible' }, this);
    }
}