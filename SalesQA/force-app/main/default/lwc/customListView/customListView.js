import { LightningElement, track, api, wire } from 'lwc';
import { NavigationMixin } from 'lightning/navigation';
import getMasterOrderRecords from '@salesforce/apex/CustomMasterOrderListViewController.getObjectRecords';
import getObjectRecords from '@salesforce/apex/CustomListViewController.getObjectRecords'; 
import buscar from '@salesforce/apex/CustomOrderListViewController.buscar';
import basePath from '@salesforce/community/basePath';


/**
 * @slot Extra-Space
*/


export default class B2BObjectTable extends NavigationMixin(LightningElement) {
    @api objeto;
    @api titulo;
    @api campos;
    @api formato;
    @api filtro;
    @api enablequickactions;
    @api recordPageName;
    @api iconName;
    @api emptyMessage;
    @api sortField;
    @api sortDirect;
    @api nomeColunaAcao= '';
    @api enableDetailPage;
    @api detailFields;
    @api detailFieldsSources
    @api detailFieldsTypes
    @api showPrintButton
    @api showLoadMore

    @track data = [];
    @track columns = [];
    @track actions = [];
    @track listTitulo = [];
    @track showRecordPage = false;
    @track loaded = true;
    @track linkRecordPage = '';
    @track loadingTable = true;
    @track showEmptyText = false;
    @track isOrder = false;

    listCampos = []
    listFormato = []

    containRelation = false;
    relationFields = []

    possuiCompostoSemLink = false;

    dataMemory = [];

    defaultSortDirection = 'asc';
    sortDirection = 'asc';
    sortedBy;
    offsetNumber = 0;

    textoPesquisa = ''

    getRecordsFunciont


    listCamposDetalhes
    listCamposDetalhesOrigem
    listCamposDetalhesTipo

    lstCamposPicklistLabels = []

    async connectedCallback() {

        
        // define qual funçao será utilizada
        if(this.objeto == 'MasterOrderItem__c' || this.objeto == 'MasterOrder__c'){
            this.getRecordsFunciont = getMasterOrderRecords;
            this.isOrder = true;
        }
        else{
            this.getRecordsFunciont = getObjectRecords;
        }

        if(this.objeto == 'OrderItem' ){

        }
        console.log(this.objeto);
        
        

        if (this.titulo && this.campos){
            // separa as entradas em string em listas para melhor manipulação
            this.listTitulo = this.titulo.split(',');
            this.listCampos = this.campos.split(',');
            this.listFormato = this.formato.split(',');
            this.columns = [];
            // Caso páginas de detalhe estejam ativada (links)
            if(this.detailFields && this.enableDetailPage){
                // separa as variaveis em string necessáris em listas
                this.listCamposDetalhes = this.detailFields.split(',');
                this.listCamposDetalhesOrigem = this.detailFieldsSources.split(',');
                this.listCamposDetalhesTipo= this.detailFieldsTypes.split(',');

                // determina as caracteristiacas da tabela
                for (let n = 0; n < this.listTitulo.length; n++) {
                    // converte todo campo composto ex: 'Order.Id' para um nome simples, substituindo o '.' para que possa ser exibido na tabela
                    let campo = this.listCampos[n].replace(/\./g, '->');
                    
                    // to get picklist label
                    if(campo.includes('toLabel')){
                        campo = campo.replace('toLabel(', '').replace(')','');
                    }
                    

                    // verifica se a coluna tera um link
                    if((this.listCamposDetalhes.includes(this.listCampos[n]) && this.enableDetailPage)){ 
                        // tipo do campo definido como url                      
                        this.listFormato[n] = 'url';
                        // fieldName: nome do campo que vai ter o url, typeAttributes.fieldName: nome do campo que aparece para o usuario
                        this.columns.push({ label: this.listTitulo[n], fieldName: 'DetailPage' + campo, type: this.listFormato[n], typeAttributes:{label: {fieldName: campo}}, sortable: true, wrapText: true});
                    }
                    else{
                        if(campo.includes('->')){
                            this.possuiCompostoSemLink = true;
                        }
                        if(this.listFormato[n] === 'number6'){
                            this.columns.push({ label: this.listTitulo[n], fieldName: campo, type: 'number',typeAttributes:{maximumFractionDigits: 6, minimumFractionDigits: 6} ,sortable: true, wrapText: true });
                        }
                        else if(this.listFormato[n] === 'date' || this.listFormato[n] === 'date-local'){
                            this.columns.push({ label: this.listTitulo[n], fieldName: campo, type: this.listFormato[n] ,typeAttributes:{timeZone:'UTC'} ,sortable: true, wrapText: true });
                        }
                        else{
                            this.columns.push({ label: this.listTitulo[n], fieldName: campo, type: this.listFormato[n], sortable: true, wrapText: true });
                        }
                    }
                }
            }
            // Caso padrão preenche de forma padrão
            else{   
                for (let n = 0; n < this.listTitulo.length; n++) {
                    
                    // converte todo campo composto ex: 'Order.Id' para um nome simples, substituindo o '.' para que possa ser exibido na tabela
                    let campo = this.listCampos[n].replace(/\./g, '->');

                    // to get picklist label
                    if(campo.includes('toLabel')){
                        campo = campo.replace('toLabel(', '').replace(')','');
                    }


                    if(campo.includes('->')){
                        this.possuiCompostoSemLink = true;
                    }
                    this.columns.push({ label: this.listTitulo[n], fieldName: campo, type: this.listFormato[n], sortable: true, wrapText: true });
                }
            }
            if (this.enablequickactions) {
                this.columns.push({label: this.nomeColunaAcao, type: 'button-icon', typeAttributes:{iconName: this.iconName, name: 'edit'}, wrapText: true});
            
            }
            // trata os campos com os necessarios para obter os ids para que se encaixem na query
            if(!(this.detailFieldsSources === undefined)){
                this.detailFieldsSources = this.detailFieldsSources.replace(/(,Id)|(^Id)(?!.Id,)/g,'');
                if(this.detailFieldsSources == null || this.detailFieldsSources == ''){
                    this.detailFieldsSources = '';
                }
                else{
                    this.detailFieldsSources = ',' + this.detailFieldsSources;
                }
            }else{
                this.detailFieldsSources = '';
            }

            this.checkIfFilterContainsUndefinedAndGetData();
        }
    }

    getData(){
        this.getRecordsFunciont({filter: this.filtro, fields: this.campos + this.detailFieldsSources, objectName: this.objeto, sortField: this.sortField, sortDirect: this.sortDirect})
                .then(result => {
                    this.loadingTable = false;
                    if (result.length > 0) {
                        // caso existam campos compostos ou referencias
                        if(this.enableDetailPage || this.possuiCompostoSemLink){
                            let indexPath = window.location.pathname.indexOf("s/") + 2;
                            let basePathName = basePath + '/';  //window.location.pathname.slice(0, indexPath);
                            var tmpRecords = [];
                            var tmpData = result;
                            tmpData.forEach( (record) =>{
                                let tmpRecord = Object.assign({}, record);
                                var i = 0;
                                var campoOld = '';
                                this.listCampos.forEach( (campo) =>{
                                    campoOld = campo;
                                    // se campo é composto faz o tratamento
                                    if(campo.includes('.')){
                                        // converte campo composto para campo unico
                                        let referencesList = campo.split('.');
                                        let acessingRecord = tmpRecord[referencesList[0]];

                                        for(let j = 1; j<referencesList.length; j++){
                                            acessingRecord = acessingRecord?.[referencesList[j]];
                                        }
                                        // converte nome 
                                        campoOld = campo;
                                        campo = campo.replace(/\./g, '->');
                                        // salva valor 
                                        tmpRecord[campo] = acessingRecord;
                                    }
                                    // caso tenha referencia popula os links
                                    if(this.enableDetailPage){
                                        if((tmpRecord[campo]) && campoOld == this.listCamposDetalhes[i]){
                                            
                                            // OrderSummary page does not need '/detail'
                                            let detail = '/detail';
                                            if(this.listCamposDetalhesTipo[i] == 'OrderSummary'){
                                                detail = '';
                                            }

                                            // se campo de Id é composto
                                            if(this.listCamposDetalhesOrigem[i].includes('.')){
                                                // converte campo de Id composto para campo unico
                                                let referencesField = this.listCamposDetalhesOrigem[i].split('.');
                                                let acessingId = tmpRecord[referencesField[0]];
                                                for(let k = 1; k<referencesField.length; k++){
                                                    acessingId = acessingId?.[referencesField[k]];
                                                }
                                                // converte nome
                                                campo = campo.replace(/\./g, '->');
                                                // aplica valor                                                
                                                
                                                tmpRecord['DetailPage' + campo] = basePathName + this.listCamposDetalhesTipo[i] + '/' + acessingId + detail;
                                            }else{
                                                tmpRecord['DetailPage' + campo] = basePathName + this.listCamposDetalhesTipo[i] + '/' + tmpRecord[this.listCamposDetalhesOrigem[i]] + detail;
                                            }
                                            i++;
                                        }
                                    }
                                } )
                                tmpRecords.push(tmpRecord);
                            })
                            this.data = tmpRecords;
                        }
                        else{
                            this.data = result;
                        }                        
                        this.data.forEach( (record) =>{
                            for(let i = 0; i<this.listCampos.length; i++){
                                var tmpCampo = this.listCampos[i].replace(/\./g, '->');
                                if(typeof(record[tmpCampo]) === 'boolean' && this.listFormato[i] == 'text'){
                                    if(record[tmpCampo]){
                                        record[tmpCampo] = 'Sim';
                                    }
                                    else{
                                        record[tmpCampo] = 'Não';
                                    }
                                }
                                /*
                                if(tmpCampo.includes('toLabel(')){
                                    tmpCampo = tmpCampo.replace('toLabel(','').replace(')','');
                                    record[tmpCampo] = this.listCampos[i];
                                }
                                    */
                            }
                        })
                        if (this.sortField && this.sortField != "" && (this.sortDirect === 'asc' || this.sortDirect === 'desc') && this.campos.indexOf(this.sortField) >= 0) {
                            const cloneData = [...this.data];
                            cloneData.sort(this.sortBy(this.sortField, this.sortDirect === 'asc' ? 1 : -1));
                            this.data = cloneData;
                            this.sortDirection = this.sortDirect;
                            this.sortedBy = this.sortField;
                        }
                    } else {
                        this.showEmptyText = true;
                    }
                })
                .catch(error => {
                    console.log('Erro: '+error.message);
                    console.log(error);
                    this.showEmptyText = true;
                });
    }
   
    stopSpinner(event) {
        this.loaded = true;
    }

    closeRecordPage(event) {
        this.showRecordPage = false;
        this.loaded = true;
    }

    // Used to sort the 'Age' column
    sortBy(isNumeric, field, reverse, primer) {
        const key = primer
            ? function (x) {
                  return primer(x[field]);
              }
            : function (x) {
                  return x[field];
              };

        return function (a, b) {
            var x
            var y
            if(isNumeric){
                x = parseFloat(key(a))
                y = parseFloat(key(b))
            }
            else{
                x = key(a);
                y = key(b);
            }
            return reverse * ((x > y) - (y > x));
        };
    }

    onHandleSort(event) {        
        let { fieldName: sortedBy, sortDirection } = event.detail;
        const cloneData = [...this.data];
        var isNumeric = false;

        var sortedByAux = sortedBy.replace('DetailPage','');

        for(let i=0; i<this.listCampos.length; i++){
            if(this.listCampos[i] == sortedByAux){
                if(this.listFormato[i] == 'number'){
                    isNumeric = true;
                }
            }
        }
 
        cloneData.sort(this.sortBy(isNumeric, sortedByAux, sortDirection === 'asc' ? 1 : -1));
        this.data = cloneData;
        this.sortDirection = sortDirection;
        this.sortedBy = sortedBy;
    }

    handlePesquisaValue(event){
        this.textoPesquisa = event.detail.value;
        if(this.textoPesquisa == '' || this.textoPesquisa == null){
            this.reloadTable();
        }
    }
    
    handleChangeFilter(event) {
        this.showLoadMore = false;
        console.log('this.textoPesquisa:'+this.textoPesquisa);
        if(event.target.value && !this.isOrder){
            this.textoPesquisa = event.target.value;
        }
        if (this.textoPesquisa) {
            if(this.objeto == 'Order'){
                buscar({filter: this.textoPesquisa, fields: this.campos + this.detailFieldsSources}).then(result => {
                    if (result.length > 0) {
                        this.data = [];
                        // caso existam campos compostos ou referencias
                        if(this.enableDetailPage || this.possuiCompostoSemLink){
                            let indexPath = window.location.pathname.indexOf("s/") + 2;
                            let basePathName = window.location.pathname.slice(0, indexPath);
                            var tmpRecords = [];
                            var tmpData = result;
                            tmpData.forEach( (record) =>{
                                let tmpRecord = Object.assign({}, record);
                                var i = 0;
                                var campoOld = '';
                                this.listCampos.forEach( (campo) =>{
                                    campoOld = campo;
                                    // se campo é composto faz o tratamento
                                    if(campo.includes('.')){
                                        // converte campo composto para campo unico
                                        let referencesList = campo.split('.');
                                        let acessingRecord = tmpRecord[referencesList[0]];

                                        for(let j = 1; j<referencesList.length; j++){
                                            acessingRecord = acessingRecord?.[referencesList[j]];
                                        }
                                        // converte nome 
                                        campoOld = campo;
                                        campo = campo.replace(/\./g, '->');
                                        // salva valor 
                                        tmpRecord[campo] = acessingRecord;
                                    }
                                    // caso tenha referencia popula os links
                                    if(this.enableDetailPage){
                                        if((tmpRecord[campo]) && campoOld == this.listCamposDetalhes[i]){
                                            // se campo de Id é composto
                                            if(this.listCamposDetalhesOrigem[i].includes('.')){
                                                // converte campo de Id composto para campo unico
                                                let referencesField = this.listCamposDetalhesOrigem[i].split('.');
                                                let acessingId = tmpRecord[referencesField[0]];
                                                for(let k = 1; k<referencesField.length; k++){
                                                    acessingId = acessingId?.[referencesField[k]];
                                                }
                                                // converte nome
                                                campo = campo.replace(/\./g, '->');
                                                // aplica valor
                                                tmpRecord['DetailPage' + campo] = basePathName + this.listCamposDetalhesTipo[i] + '/' + acessingId + '/detail';
                                            }else{
                                                tmpRecord['DetailPage' + campo] = basePathName + this.listCamposDetalhesTipo[i] + '/' + tmpRecord[this.listCamposDetalhesOrigem[i]] + '/detail';
                                            }
                                            i++;
                                        }
                                    }
                                } )
                                tmpRecords.push(tmpRecord);
                            })
                            //this.data = tmpRecords;
                            this.data = [...this.data, ...tmpRecords]
                        }
                        else{
                            this.data = [...this.data, ...result]
                            //this.data = result;
                        }                        
                        this.data.forEach( (record) =>{

                            for(let i = 0; i<this.listCampos.length; i++){
                                var tmpCampo = this.listCampos[i].replace(/\./g, '->')
                                if(typeof(record[tmpCampo]) === 'boolean' && this.listFormato[i] == 'text'){
                                    if(record[tmpCampo]){
                                        record[tmpCampo] = 'Sim';
                                    }
                                    else{
                                        record[tmpCampo] = 'Não';
                                    }
                                }
                            }
                        })
                        if (this.sortField && this.sortField != "" && (this.sortDirect === 'asc' || this.sortDirect === 'desc') && this.campos.indexOf(this.sortField) >= 0) {
                            const cloneData = [...this.data];
                            cloneData.sort(this.sortBy(this.sortField, this.sortDirect === 'asc' ? 1 : -1));
                            this.data = cloneData;
                            this.sortDirection = this.sortDirect;
                            this.sortedBy = this.sortField;
                        }

                    }
                    this.loadingTable = false;
                    this.dataMemory = [...this.data];
                    var listCampos = this.campos.replaceAll('.', '->').split(',');
                    this.data = this.dataMemory.filter(item => this.checkIfIncludes(item, listCampos, this.textoPesquisa));
                })
                .catch(error => {
                    console.log('Erro: '+error.message);
                    console.log(error);
                    this.showEmptyText = true;
                });
            }else{
                if (this.dataMemory.length === 0) this.dataMemory = [...this.data];
                var listCampos = this.campos.replaceAll('.', '->').split(',');
                this.data = this.dataMemory.filter(item => this.checkIfIncludes(item, listCampos, this.textoPesquisa));
            }
        } else {
            console.log('valor vazio');
            this.textoPesquisa = ''
            this.reloadTable();
            this.dataMemory = [];
        }
        
    }

    reloadTable(){
        this.showLoadMore = true;
        console.log('reload');
        this.getRecordsFunciont({filter: this.filtro, fields: this.campos + this.detailFieldsSources, objectName: this.objeto, sortField: this.sortField, sortDirect: this.sortDirect})
                .then(result => {
                    this.loadingTable = false;
                    if (result.length > 0) {
                        // caso existam campos compostos ou referencias
                        if(this.enableDetailPage || this.possuiCompostoSemLink){
                            let indexPath = window.location.pathname.indexOf("s/") + 2;
                            let basePathName = window.location.pathname.slice(0, indexPath);
                            var tmpRecords = [];
                            var tmpData = result;
                            tmpData.forEach( (record) =>{
                                let tmpRecord = Object.assign({}, record);
                                var i = 0;
                                var campoOld = '';
                                this.listCampos.forEach( (campo) =>{
                                    campoOld = campo;
                                    // se campo é composto faz o tratamento
                                    if(campo.includes('.')){
                                        // converte campo composto para campo unico
                                        let referencesList = campo.split('.');
                                        let acessingRecord = tmpRecord[referencesList[0]];

                                        for(let j = 1; j<referencesList.length; j++){
                                            acessingRecord = acessingRecord?.[referencesList[j]];
                                        }
                                        // converte nome 
                                        campoOld = campo;
                                        campo = campo.replace(/\./g, '->');
                                        // salva valor 
                                        tmpRecord[campo] = acessingRecord;
                                    }
                                    // caso tenha referencia popula os links
                                    if(this.enableDetailPage){
                                        if((tmpRecord[campo]) && campoOld == this.listCamposDetalhes[i]){
                                            // se campo de Id é composto
                                            if(this.listCamposDetalhesOrigem[i].includes('.')){
                                                // converte campo de Id composto para campo unico
                                                let referencesField = this.listCamposDetalhesOrigem[i].split('.');
                                                let acessingId = tmpRecord[referencesField[0]];
                                                for(let k = 1; k<referencesField.length; k++){
                                                    acessingId = acessingId?.[referencesField[k]];
                                                }
                                                // converte nome
                                                campo = campo.replace(/\./g, '->');
                                                // aplica valor
                                                tmpRecord['DetailPage' + campo] = basePathName + this.listCamposDetalhesTipo[i] + '/' + acessingId + '/detail';
                                            }else{
                                                tmpRecord['DetailPage' + campo] = basePathName + this.listCamposDetalhesTipo[i] + '/' + tmpRecord[this.listCamposDetalhesOrigem[i]] + '/detail';
                                            }
                                            i++;
                                        }
                                    }
                                } )
                                tmpRecords.push(tmpRecord);
                            })
                            this.data = tmpRecords;
                        }
                        else{
                            this.data = result;
                        }                        
                        this.data.forEach( (record) =>{

                            for(let i = 0; i<this.listCampos.length; i++){
                                var tmpCampo = this.listCampos[i].replace(/\./g, '->')
                                if(typeof(record[tmpCampo]) === 'boolean' && this.listFormato[i] == 'text'){
                                    if(record[tmpCampo]){
                                        record[tmpCampo] = 'Sim';
                                    }
                                    else{
                                        record[tmpCampo] = 'Não';
                                    }
                                }
                            }
                        })
                        if (this.sortField && this.sortField != "" && (this.sortDirect === 'asc' || this.sortDirect === 'desc') && this.campos.indexOf(this.sortField) >= 0) {
                            const cloneData = [...this.data];
                            cloneData.sort(this.sortBy(this.sortField, this.sortDirect === 'asc' ? 1 : -1));
                            this.data = cloneData;
                            this.sortDirection = this.sortDirect;
                            this.sortedBy = this.sortField;
                        }
                    } else {
                        this.showEmptyText = true;
                    }
                })
                .catch(error => {
                    console.log('Erro: '+error.message);
                    console.log(error);
                    this.showEmptyText = true;
                });
    }

    checkIfIncludes(item, listCampos, value){
        for(let i = 0; i<listCampos.length; i++){
            if(String(item[String(listCampos[i])]).toLowerCase().includes(value.toLowerCase())){
                return true
            }
        }
        return false;
    }


    handleLoadMore(item, listCampos, value){
        this.offsetNumber = this.offsetNumber + 50;
        let offsetText = this.offsetNumber.toString();
        this.loadingTable = true;
        this.getRecordsFunciont({filter: this.filtro, fields: this.campos + this.detailFieldsSources, objectName: this.objeto, sortField: this.sortField, sortDirect: this.sortDirect, offset: offsetText})
                .then(result => {
                    this.loadingTable = false;
                    if (result.length > 0) {
                        // caso existam campos compostos ou referencias
                        if(this.enableDetailPage || this.possuiCompostoSemLink){
                            let indexPath = window.location.pathname.indexOf("s/") + 2;
                            let basePathName = window.location.pathname.slice(0, indexPath);
                            var tmpRecords = [];
                            var tmpData = result;
                            tmpData.forEach( (record) =>{
                                let tmpRecord = Object.assign({}, record);
                                var i = 0;
                                var campoOld = '';
                                this.listCampos.forEach( (campo) =>{
                                    campoOld = campo;
                                    // se campo é composto faz o tratamento
                                    if(campo.includes('.')){
                                        // converte campo composto para campo unico
                                        let referencesList = campo.split('.');
                                        let acessingRecord = tmpRecord[referencesList[0]];

                                        for(let j = 1; j<referencesList.length; j++){
                                            acessingRecord = acessingRecord?.[referencesList[j]];
                                        }
                                        // converte nome 
                                        campoOld = campo;
                                        campo = campo.replace(/\./g, '->');
                                        // salva valor 
                                        tmpRecord[campo] = acessingRecord;
                                    }
                                    // caso tenha referencia popula os links
                                    if(this.enableDetailPage){
                                        if((tmpRecord[campo]) && campoOld == this.listCamposDetalhes[i]){
                                            // se campo de Id é composto
                                            if(this.listCamposDetalhesOrigem[i].includes('.')){
                                                // converte campo de Id composto para campo unico
                                                let referencesField = this.listCamposDetalhesOrigem[i].split('.');
                                                let acessingId = tmpRecord[referencesField[0]];
                                                for(let k = 1; k<referencesField.length; k++){
                                                    acessingId = acessingId?.[referencesField[k]];
                                                }
                                                // converte nome
                                                campo = campo.replace(/\./g, '->');
                                                // aplica valor
                                                tmpRecord['DetailPage' + campo] = basePathName + this.listCamposDetalhesTipo[i] + '/' + acessingId + '/detail';
                                            }else{
                                                tmpRecord['DetailPage' + campo] = basePathName + this.listCamposDetalhesTipo[i] + '/' + tmpRecord[this.listCamposDetalhesOrigem[i]] + '/detail';
                                            }
                                            i++;
                                        }
                                    }
                                } )
                                tmpRecords.push(tmpRecord);
                            })
                            //this.data = tmpRecords;
                            this.data = [...this.data, ...tmpRecords]
                        }
                        else{
                            this.data = [...this.data, ...result]
                            //this.data = result;
                        }                        
                        this.data.forEach( (record) =>{

                            for(let i = 0; i<this.listCampos.length; i++){
                                var tmpCampo = this.listCampos[i].replace(/\./g, '->')
                                if(typeof(record[tmpCampo]) === 'boolean' && this.listFormato[i] == 'text'){
                                    if(record[tmpCampo]){
                                        record[tmpCampo] = 'Sim';
                                    }
                                    else{
                                        record[tmpCampo] = 'Não';
                                    }
                                }
                            }
                        })
                        if (this.sortField && this.sortField != "" && (this.sortDirect === 'asc' || this.sortDirect === 'desc') && this.campos.indexOf(this.sortField) >= 0) {
                            const cloneData = [...this.data];
                            cloneData.sort(this.sortBy(this.sortField, this.sortDirect === 'asc' ? 1 : -1));
                            this.data = cloneData;
                            this.sortDirection = this.sortDirect;
                            this.sortedBy = this.sortField;
                        }
                    } else {

                        this.showLoadMore = false
                    }
                })
                .catch(error => {
                    console.log('Erro: '+error.message);
                    console.log(error);
                    this.showEmptyText = true;
                });
    }

    async checkIfFilterContainsUndefinedAndGetData(){
        setTimeout(() => {
            console.log('Verificando undefined:', this.filtro);
            if(this.filtro.includes('undefined')){
                this.checkIfFilterContainsUndefinedAndGetData();
            }
            else{
                this.getData();
            }
        }, 50);
    }
}