import { LightningElement, track } from 'lwc';
import { getSessionContext } from 'commerce/contextApi';
import buscarCodigoERPApex   from '@salesforce/apex/BoletosController.buscarCodigoERP';
import buscarBoletosClienteApex     from '@salesforce/apexContinuation/BoletosController.buscarBoletosCliente';
import buscarNotasApex       from '@salesforce/apexContinuation/BoletosController.buscarNotas';
import downloadDocumentoApex from '@salesforce/apexContinuation/BoletosController.downloadDocumento';

export default class BoletosClienteLWC extends LightningElement {

    @track boletos   = [];
    @track notas     = [];
    @track abaAtiva  = 'boletos';

    @track isLoadingBoletos = false;
    @track isLoadingNotas   = false;
    @track errorBoletos     = '';
    @track errorNotas       = '';
    @track accountName      = '';

    effectiveAccountId = '';
    codigoERP          = '';

    async connectedCallback() {
        try {
            const context = await getSessionContext();
            this.effectiveAccountId = context?.effectiveAccountId ?? '';
            this.accountName        = context?.effectiveAccountName ?? '';

            if (this.effectiveAccountId) {
                const codigoERP = await buscarCodigoERPApex({
                    accountId: this.effectiveAccountId
                });
                this.codigoERP = codigoERP ?? '';
                console.log('Código ERP:', this.codigoERP);
            }
        } catch (error) {
            console.error('Erro ao carregar contexto:', error);
            this.errorBoletos = 'Não foi possível carregar o contexto do cliente.';
        }
    }

    // ── Abas ──────────────────────────────────────────

    get abaBoletosAtiva() { return this.abaAtiva === 'boletos'; }
    get abaNotasAtiva()   { return this.abaAtiva === 'notas'; }

    get classAbaBoletos() {
        return this.abaAtiva === 'boletos'
            ? 'slds-tabs_default__item slds-is-active'
            : 'slds-tabs_default__item';
    }

    get classAbaNotas() {
        return this.abaAtiva === 'notas'
            ? 'slds-tabs_default__item slds-is-active'
            : 'slds-tabs_default__item';
    }

    handleAba(event) {
        this.abaAtiva     = event.currentTarget.dataset.aba;
        this.errorBoletos = '';
        this.errorNotas   = '';
    }

    // ── Boletos ───────────────────────────────────────

    async buscarBoletos() {
        if (!this.codigoERP) {
            this.errorBoletos = 'Código ERP não encontrado.';
            return;
        }

        this.isLoadingBoletos = true;
        this.errorBoletos     = '';
        this.boletos          = [];

        try {
            const resultado = await buscarBoletosClienteApex({ codigoERP: this.codigoERP });
            const parsed    = JSON.parse(resultado);
            const lista     = parsed?.boletos ?? [];

            if (lista.length === 0) {
                this.errorBoletos = 'Nenhum boleto encontrado.';
                return;
            }

            this.boletos = lista.map((b, i) => ({
                ...b,
                index:       i,
                selecionado: false,
                baixando:    false
            }));

        } catch (error) {
            console.error('Erro boletos:', error);
            this.errorBoletos = error?.body?.message ?? error?.message ?? 'Erro desconhecido.';
        } finally {
            this.isLoadingBoletos = false;
        }
    }

    handleCheckboxBoleto(event) {
        const index   = parseInt(event.currentTarget.dataset.index, 10);
        const checked = event.target.checked;
        this.boletos  = this.boletos.map((b, i) => ({
            ...b,
            selecionado: i === index ? checked : b.selecionado
        }));
    }

    handleSelecionarTodosBoletos(event) {
        const checked = event.target.checked;
        this.boletos  = this.boletos.map(b => ({ ...b, selecionado: checked }));
    }

    async handleDownloadBoletoUnico(event) {
        const index  = parseInt(event.currentTarget.dataset.index, 10);
        const boleto = this.boletos[index];
        await this.executarDownload('boleto', boleto.arquivo, index, 'boletos');
    }

    async handleDownloadBoletosSelecionados() {
        const selecionados = this.boletos
            .map((b, i) => ({ ...b, index: i }))
            .filter(b => b.selecionado);

        if (selecionados.length === 0) {
            this.errorBoletos = 'Selecione ao menos um boleto.';
            return;
        }

        this.errorBoletos = '';
        for (const b of selecionados) {
            await this.executarDownload('boleto', b.arquivo, b.index, 'boletos');
        }
    }

    get temBoletos()            { return this.boletos.length > 0; }
    get temBoletosSelecionados(){ return this.boletos.some(b => b.selecionado); }
    get totalBoletosSelecionados() {
        return this.boletos.filter(b => b.selecionado).length;
    }
    get labelBaixarBoletos() {
        const t = this.totalBoletosSelecionados;
        return t === 1 ? 'Baixar 1 boleto' : `Baixar ${t} boletos`;
    }

    // ── Notas Fiscais ─────────────────────────────────

    async buscarNotas() {
        if (!this.codigoERP) {
            this.errorNotas = 'Código ERP não encontrado.';
            return;
        }

        this.isLoadingNotas = true;
        this.errorNotas     = '';
        this.notas          = [];

        try {
            const resultado = await buscarNotasApex({ codigoERP: this.codigoERP });
            const parsed    = JSON.parse(resultado);
            const lista     = parsed?.notas ?? [];

            if (lista.length === 0) {
                this.errorNotas = 'Nenhuma nota fiscal encontrada.';
                return;
            }

            this.notas = lista.map((n, i) => ({
                ...n,
                index:       i,
                selecionado: false,
                baixando:    false
            }));

        } catch (error) {
            console.error('Erro notas:', error);
            this.errorNotas = error?.body?.message ?? error?.message ?? 'Erro desconhecido.';
        } finally {
            this.isLoadingNotas = false;
        }
    }

    handleCheckboxNota(event) {
        const index   = parseInt(event.currentTarget.dataset.index, 10);
        const checked = event.target.checked;
        this.notas    = this.notas.map((n, i) => ({
            ...n,
            selecionado: i === index ? checked : n.selecionado
        }));
    }

    handleSelecionarTodasNotas(event) {
        const checked = event.target.checked;
        this.notas    = this.notas.map(n => ({ ...n, selecionado: checked }));
    }

    async handleDownloadNotaUnica(event) {
        const index = parseInt(event.currentTarget.dataset.index, 10);
        const nota  = this.notas[index];
        await this.executarDownload('nf', nota.nomeArquivoNF, index, 'notas');
    }

    async handleDownloadNotasSelecionadas() {
        const selecionadas = this.notas
            .map((n, i) => ({ ...n, index: i }))
            .filter(n => n.selecionado);

        if (selecionadas.length === 0) {
            this.errorNotas = 'Selecione ao menos uma nota.';
            return;
        }

        this.errorNotas = '';
        for (const n of selecionadas) {
            await this.executarDownload('nf', n.nomeArquivoNF, n.index, 'notas');
        }
    }

    get temNotas()            { return this.notas.length > 0; }
    get temNotasSelecionadas(){ return this.notas.some(n => n.selecionado); }
    get totalNotasSelecionadas() {
        return this.notas.filter(n => n.selecionado).length;
    }
    get labelBaixarNotas() {
        const t = this.totalNotasSelecionadas;
        return t === 1 ? 'Baixar 1 nota fiscal' : `Baixar ${t} notas fiscais`;
    }

    // ── Download ──────────────────────────────────────

    async executarDownload(tipo, nomeArquivo, index, lista) {
        // Marca como baixando
        this[lista] = this[lista].map((item, i) => ({
            ...item,
            baixando: i === index ? true : item.baixando
        }));

        try {
            const resultado = await downloadDocumentoApex({ tipo, nomeArquivo });
            const parsed    = JSON.parse(resultado);

            if (!parsed.sucesso || !parsed.pdf) {
                const erro = 'Arquivo não disponível: ' + nomeArquivo;
                if (lista === 'boletos') this.errorBoletos = erro;
                else                     this.errorNotas   = erro;
                return;
            }

            this.dispararDownload(parsed.pdf, nomeArquivo);

        } catch (error) {
            console.error('Erro download:', error);
            const msg = error?.body?.message ?? 'Erro ao baixar arquivo.';
            if (lista === 'boletos') this.errorBoletos = msg;
            else                     this.errorNotas   = msg;
        } finally {
            this[lista] = this[lista].map((item, i) => ({
                ...item,
                baixando: i === index ? false : item.baixando
            }));
        }
    }

    dispararDownload(base64, nomeArquivo) {
        try {
            const byteCharacters = atob(base64);
            const byteNumbers    = new Array(byteCharacters.length);
            for (let i = 0; i < byteCharacters.length; i++) {
                byteNumbers[i] = byteCharacters.charCodeAt(i);
            }
            const byteArray = new Uint8Array(byteNumbers);
            const blob      = new Blob([byteArray], { type: 'application/pdf' });
            const url       = URL.createObjectURL(blob);
            const link      = document.createElement('a');
            link.href       = url;
            link.download   = nomeArquivo;
            link.click();
            URL.revokeObjectURL(url);
        } catch (error) {
            console.error('Erro ao disparar download:', error);
        }
    }
}