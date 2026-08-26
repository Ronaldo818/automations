import { LightningElement, track } from 'lwc';
import { getSessionContext } from 'commerce/contextApi';
import buscarCodigoERPApex from '@salesforce/apex/BoletosController.buscarCodigoERP';
import buscarBoletosApex from '@salesforce/apexContinuation/BoletosController.buscarBoletos';
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

    columns = [
        { label: 'Referência', fieldName: 'referencia' },
        { label: 'Número',     fieldName: 'numero' },
        { label: 'Vencimento', fieldName: 'vencimento' },
        { label: 'Valor (R$)', fieldName: 'valor' },
        { label: 'Observação', fieldName: 'observacao' }
    ];

    async connectedCallback() {
        try {
            const context = await getSessionContext();
            console.log('Session context:', JSON.stringify(context));

            this.effectiveAccountId = context?.effectiveAccountId ?? '';
            this.accountName        = context?.effectiveAccountName ?? '';

            console.log('effectiveAccountId:', this.effectiveAccountId);

            // Busca o código ERP direto do campo da Account
            if (this.effectiveAccountId) {
                const codigoERP = await buscarCodigoERPApex({
                    accountId: this.effectiveAccountId
                });
                this.codigoERP = codigoERP ?? '';
                console.log('Código ERP (campo):', this.codigoERP);
            }

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

        this.isLoading              = true;
        this.errorMessage           = '';
        this.successMessage         = '';
        this.boletos                = [];
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

            // Mapeia os dados e cria as novas propriedades para controle visual
            this.boletos = lista.map((b, i) => {
                const semPdf = !b.pdf || b.pdf.trim() === '';
                
                return {
                    ...b,
                    index:       i,
                    selecionado: false,
                    semPdf:      semPdf,
                    // Aqui inserimos o \n para forçar a quebra na tela
                    observacao:  semPdf ? 'Arquivo não disponível.\nFavor solicitar ao setor financeiro.' : '',
                    cssClass:    semPdf ? 'slds-hint-parent slds-theme_shade' : 'slds-hint-parent',
                    cursorStyle: semPdf ? 'cursor: not-allowed; opacity: 0.6;' : 'cursor: pointer;'
                };
            });

        } catch (error) {
            console.error('Erro Apex:', error);
            this.errorMessage = error?.body?.message ?? error?.message ?? 'Erro desconhecido.';
        } finally {
            this.isLoading = false;
        }
    }

    handleSelecionarBoleto(event) {
        const index = parseInt(event.currentTarget.dataset.index, 10);
        
        // Se a linha clicada não tiver PDF, interrompe a ação e não seleciona
        if (this.boletos[index].semPdf) {
            return;
        }

        this.boletos = this.boletos.map((b, i) => ({
            ...b,
            selecionado: i === index,
            cssClass:    i === index ? 'slds-is-selected slds-hint-parent' : (b.semPdf ? 'slds-hint-parent slds-theme_shade' : 'slds-hint-parent')
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
                nomeArquivo: boleto.arquivo,
                valor:       String(boleto.valor),
                vencimento:  String(boleto.vencimento),
                referencia:  String(boleto.referencia)
            });

            this.successMessage = 'Boleto enviado com sucesso para o email do cliente!';

            // Limpa a seleção e mantém o visual de bloqueio nos itens sem pdf
            this.boletos = this.boletos.map(b => ({
                ...b,
                selecionado: false,
                cssClass:    b.semPdf ? 'slds-hint-parent slds-theme_shade' : 'slds-hint-parent'
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