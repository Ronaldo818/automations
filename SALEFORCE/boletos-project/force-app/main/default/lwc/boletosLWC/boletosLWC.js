import { LightningElement, track } from 'lwc';
import { getSessionContext } from 'commerce/contextApi';
import buscarBoletosApex from '@salesforce/apex/BoletosController.buscarBoletos';
import enviarBoletoEmailApex from '@salesforce/apex/BoletosController.enviarBoletoEmail';

export default class BoletosLWC extends LightningElement {

    @track boletos = [];
    @track isLoading = false;
    @track isSending = false;
    @track errorMessage = '';
    @track successMessage = '';
    @track accountName = '';

    effectiveAccountId = '';
    codigoERP = '';
    boletoselecionadoIndex = null;

    async connectedCallback() {
        try {
            const context = await getSessionContext();
            console.log('Session context:', JSON.stringify(context));

            this.effectiveAccountId = context?.effectiveAccountId ?? '';
            this.accountName        = context?.effectiveAccountName ?? '';

            const partes = this.accountName.split(' - ');
            this.codigoERP = partes[0]?.trim() ?? '';

            console.log('effectiveAccountId:', this.effectiveAccountId);
            console.log('Código ERP:', this.codigoERP);

        } catch (error) {
            console.error('Erro ao carregar contexto:', error);
            this.errorMessage = 'Não foi possível carregar o contexto do cliente.';
        }
    }

    async buscarBoletos() {
        if (!this.codigoERP) {
            this.errorMessage = 'Código ERP não encontrado para este cliente.';
            return;
        }

        if (!/^\d+$/.test(this.codigoERP)) {
            this.errorMessage = 'Formato de código ERP inesperado: ' + this.codigoERP;
            return;
        }

        this.isLoading             = true;
        this.errorMessage          = '';
        this.successMessage        = '';
        this.boletos               = [];
        this.boletoselecionadoIndex = null;

        try {
            const resultado = await buscarBoletosApex({ codigoERP: this.codigoERP });
            console.log('Retorno Apex recebido');

            const parsed = JSON.parse(resultado);
            const lista  = parsed?.boletos ?? [];

            if (lista.length === 0) {
                this.errorMessage = 'Nenhum boleto em aberto encontrado.';
                return;
            }

            this.boletos = lista.map((b, i) => ({
                ...b,
                index:      i,
                selecionado: false,
                cssClass:   ''
            }));

        } catch (error) {
            console.error('Erro Apex:', error);
            this.errorMessage = error?.body?.message ?? error?.message ?? 'Erro desconhecido.';
        } finally {
            this.isLoading = false;
        }
    }

    handleSelecionarBoleto(event) {
        const index = parseInt(event.currentTarget.dataset.index, 10);

        this.boletos = this.boletos.map((b, i) => ({
            ...b,
            selecionado: i === index,
            cssClass:    i === index ? 'slds-is-selected' : ''
        }));

        this.boletoselecionadoIndex = index;
        this.successMessage = '';
        this.errorMessage   = '';
    }

    async handleEnviarEmail() {
        if (this.boletoselecionadoIndex === null) {
            this.errorMessage = 'Selecione um boleto antes de enviar.';
            return;
        }

        const boleto = this.boletos[this.boletoselecionadoIndex];

        this.isSending      = true;
        this.errorMessage   = '';
        this.successMessage = '';

        try {
            await enviarBoletoEmailApex({
                accountId:   this.effectiveAccountId,
                pdfBase64:   boleto.pdf,
                nomeArquivo: boleto.arquivo
            });

            this.successMessage = 'Boleto enviado com sucesso para o email do cliente!';

            // Limpa seleção após envio
            this.boletos = this.boletos.map(b => ({
                ...b,
                selecionado: false,
                cssClass:    ''
            }));
            this.boletoselecionadoIndex = null;

        } catch (error) {
            console.error('Erro ao enviar email:', error);
            this.errorMessage = error?.body?.message ?? error?.message ?? 'Erro ao enviar email.';
        } finally {
            this.isSending = false;
        }
    }

    get temBoletos() {
        return this.boletos.length > 0;
    }

    get temSelecionado() {
        return this.boletoselecionadoIndex !== null;
    }
}