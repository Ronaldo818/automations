import { LightningElement, api, track } from 'lwc';
import buscarDetalhamentoLogistico from '@salesforce/apex/LogisticaController.buscarDetalhamentoLogistico';

export default class LogisticaLWC extends LightningElement {
    @api recordId;
    @track isModalOpen = false;
    @track isLoading = false;
    @track errorMessage;
    @track dadosLogistica;

    abrirModal() {
        this.isModalOpen = true;
        this.buscarDadosNoN8n();
    }

    fecharModal() {
        this.isModalOpen = false;
        this.dadosLogistica = null;
        this.errorMessage = null;
    }

    buscarDadosNoN8n() {
        if (!this.recordId || this.recordId.includes('{!')) {
            this.errorMessage = 'Modo de visualização: Abra um pedido real para consultar.';
            return;
        }

        this.isLoading = true;
        this.errorMessage = null;

        buscarDetalhamentoLogistico({ recordId: this.recordId })
            .then((result) => {
                let resposta;
                
                try {
                    resposta = JSON.parse(result);
                } catch (parseError) {
                    this.errorMessage = 'A integração respondeu em um formato inesperado. Verifique os logs (F12).';
                    return;
                }

                if (resposta.hasError) {
                    this.errorMessage = resposta.errorMessage;
                } else {
                    resposta.roteirizacao = this.formatarData(resposta.roteirizacao);
                    resposta.ativacaoViagem = this.formatarData(resposta.ativacaoViagem);
                    resposta.saidaOrigem = this.formatarData(resposta.saidaOrigem);
                    resposta.chegadaCliente = this.formatarData(resposta.chegadaCliente);
                    resposta.inicioDescarregamento = this.formatarData(resposta.inicioDescarregamento);
                    resposta.fimDescarregamento = this.formatarData(resposta.fimDescarregamento);
                    
                    resposta.statusGeral = this.calcularStatusAtual(resposta);
                    
                    // ==========================================
                    // TIMELINE (5 ETAPAS)
                    // ==========================================
                    const temDado = (campo) => campo && campo !== '-' && campo.trim() !== '';

                    let s1 = temDado(resposta.roteirizacao);
                    let s2 = temDado(resposta.ativacaoViagem);
                    let s3 = temDado(resposta.saidaOrigem);
                    let s4 = temDado(resposta.chegadaCliente) || temDado(resposta.inicioDescarregamento);
                    let s5 = temDado(resposta.fimDescarregamento);

                    resposta.isRoteirizado = s1;
                    resposta.isAtivado = s2;
                    resposta.isEmTransito = s3;
                    resposta.isDescarregando = s4;
                    resposta.isEntregue = s5;

                    resposta.classS1 = s1 ? 'slds-progress__item slds-is-completed' : 'slds-progress__item slds-is-active';
                    
                    if (s2) resposta.classS2 = 'slds-progress__item slds-is-completed';
                    else if (s1) resposta.classS2 = 'slds-progress__item slds-is-active';
                    else resposta.classS2 = 'slds-progress__item';

                    if (s3) resposta.classS3 = 'slds-progress__item slds-is-completed';
                    else if (s2) resposta.classS3 = 'slds-progress__item slds-is-active';
                    else resposta.classS3 = 'slds-progress__item';

                    if (s4) resposta.classS4 = 'slds-progress__item slds-is-completed';
                    else if (s3) resposta.classS4 = 'slds-progress__item slds-is-active';
                    else resposta.classS4 = 'slds-progress__item';

                    if (s5) resposta.classS5 = 'slds-progress__item slds-is-completed';
                    else if (s4) resposta.classS5 = 'slds-progress__item slds-is-active';
                    else resposta.classS5 = 'slds-progress__item';
                    
                    this.dadosLogistica = resposta;
                }
            })
            .catch((error) => {
                this.errorMessage = 'Erro de comunicação com o servidor do Salesforce.';
            })
            .finally(() => {
                this.isLoading = false;
            });
    }

    formatarData(dataBruta) {
        if (!dataBruta || dataBruta === '-' || dataBruta.length < 10) return dataBruta;
        try {
            let partes = dataBruta.split(' ');
            let dataSplit = partes[0].split('-');
            let horaSplit = partes[1] ? partes[1].split(':') : ['00', '00'];
            return `${dataSplit[2]}/${dataSplit[1]}/${dataSplit[0]} - ${horaSplit[0]}h${horaSplit[1]}`;
        } catch (e) { return dataBruta; }
    }

    calcularStatusAtual(dados) {
        const temDado = (campo) => campo && campo !== '-' && campo.trim() !== '';
        
        // Regra ajustada: "Entregue" depende apenas do Fim do Descarregamento
        if (temDado(dados.fimDescarregamento)) return 'Entregue';
        if (temDado(dados.chegadaCliente) || temDado(dados.inicioDescarregamento)) return 'Em descarregamento';
        if (temDado(dados.saidaOrigem)) return 'Mercadoria em trânsito';
        if (temDado(dados.ativacaoViagem)) return 'Viagem ativada';
        if (temDado(dados.roteirizacao)) return 'Roteirizado';
        
        return 'Aguardando processamento';
    }
}