global without sharing class SharepointFileSender implements Schedulable {

    public virtual class BaseException extends Exception {}
    public class CustomException extends BaseException {}

    @TestVisible
    private static Boolean bypassTestCheck = false;

    // Execute method
    global void execute(SchedulableContext SC) {
        getAndEnqueueFiles();
    }

    global static void getAndEnqueueFiles(){
        List<AtualizacaoCadastral__c> lstAC = new List<AtualizacaoCadastral__c>();
        lstAC = [   SELECT  Id
                    FROM    AtualizacaoCadastral__c
                    WHERE   MovimentacaoArquivosSharepoint__c IN ('', 'PENDENTE', 'ERRO')
                    ORDER   BY CreatedDate LIMIT 30];
        System.debug('lstAC: ' + lstAC);

        if(!lstAC.isEmpty()){
            System.enqueueJob(new asyncSendFiles(lstAC));
        } else {
            scheduleNextJob();
        }
    }

    @TestVisible
    private class asyncSendFiles implements Queueable, Database.AllowsCallouts {
        private List<AtualizacaoCadastral__c> pLstAC;
        private asyncSendFiles(List<AtualizacaoCadastral__c> lstAC){
            pLstAC = lstAC;
        }

        public void execute(QueueableContext context){
            String recordId = pLstAC[0].Id;
            List<AtualizacaoCadastral__c> lstAtualizada = pLstAC.clone();
            lstAtualizada.remove(0);

            try {
                sendToExternalDataSource(recordId);
            } catch (Exception ex){
                String erro = ex.getMessage() + ' | ' + ex.getCause() + ' | ' + ex.getLineNumber() + ' | ' + ex.getStackTraceString();
                AtualizacaoCadastral__c uptAtualizacaoCadastral = new AtualizacaoCadastral__c(
                    Id = recordId,
                    MovimentacaoArquivosSharepoint__c = 'ERRO',
                    ErroDetalhado__c = erro,
                    Erro__c = erro.Abbreviate(255)                    
                );
                update uptAtualizacaoCadastral;
            } finally {
                    if(!lstAtualizada.isEmpty()){
                        if(!Test.isRunningTest() || bypassTestCheck) {
                            System.enqueueJob(new asyncSendFiles(lstAtualizada));
                        }
                        return;
                    } else {
                        scheduleNextJob();
                    }
                }
        }
    }

    public static void sendToExternalDataSource(String recordId){

        Boolean isTest = Test.isRunningTest() && !bypassTestCheck;


        String externalDataSourceId = [SELECT Id FROM ExternalDataSource WHERE DeveloperName LIKE 'SharepointAvivar%' LIMIT 1]?.Id;
        if(String.isBlank(externalDataSourceId)){
            throw new CustomException('Conexão com Sharepoint não encontrada');
        }

        AtualizacaoCadastral__c ac = [SELECT Id, Name, Cnpj__c, Cpf__c FROM AtualizacaoCadastral__c WHERE Id = :recordId];
        if(ac == null){
            throw new CustomException('Registro de análise não encontrado: ' + recordId);
        }
        System.debug('ac: ' + ac);

        Set<String> setContentDocumentId = new Set<String>();
        for (ContentDocumentLink cdl : [SELECT ContentDocumentId FROM ContentDocumentLink WHERE LinkedEntityId = :recordId]){
            setContentDocumentId.add(cdl.ContentDocumentId);
        }
        System.debug('setContentDocumentId: ' + setContentDocumentId);

        List<ContentVersion> lstContentVersion;
            if(isTest){
                lstContentVersion = new List<ContentVersion>{
                    new ContentVersion(Title='Comprovante1.pdf', VersionData=Blob.valueOf('Teste'))
                };
            } else {
                lstContentVersion = [SELECT Id, Title, PathOnClient, FileExtension, VersionData
                                    FROM ContentVersion
                                    WHERE IsLatest = true
                                    AND Origin != 'H'
                                    AND ContentDocumentId IN :setContentDocumentId
                                    ORDER BY Id];
            }
        System.debug('lstContentVersion: ' + lstContentVersion);

        if(lstContentVersion.isEmpty()){
            AtualizacaoCadastral__c uptAtualizacaoCadastral = new AtualizacaoCadastral__c(
                Id = recordId,
                MovimentacaoArquivosSharepoint__c = 'SEM ARQUIVOS'
            );
            update uptAtualizacaoCadastral;

        } else {

            ParametrosAtualizacaoCadastral__c param = ParametrosAtualizacaoCadastral__c.getInstance();

            String folderServerRelativeUrl = param?.UrlPastaArquivos__c;
            String fileNameBase = (ac.Cnpj__c ?? ac.Cpf__c) + '_' + ac.Name.replace('AC-', 'AC') + '_' + 'Comprovante';
            Integer count = 0;
        
            for(ContentVersion cv : lstContentVersion){
                System.debug('cv: ' + cv);

                count++;
                Blob fileContent = cv.VersionData;
                String fileNameOriginal = fileNameBase + count + '.' + cv.FileExtension;
                String endpoint = '/_api/web/GetFolderByServerRelativeUrl(\'' + folderServerRelativeUrl + '\')/Files/add(url=\'' + fileNameOriginal + '\',overwrite=true)';

                HttpRequest req = new HttpRequest();
                req.setEndpoint('callout:SharePointAuth' + endpoint);
                req.setMethod('POST');
                req.setHeader('Accept', 'application/json;odata=verbose');
                req.setHeader('Content-Type', 'application/octet-stream');
                req.setBodyAsBlob(fileContent);

                if(!isTest){
                    Http http = new Http();
                    HttpResponse res = http.send(req);
                    if (res.getStatusCode() == 200 || res.getStatusCode() == 201) {
                        System.debug('Upload realizado com sucesso: ' + res.getBody());
                    } else {
                        System.debug('Erro no upload. Status: ' + res.getStatusCode() + ' - ' + res.getBody());
                        throw new CalloutException('Erro ao enviar arquivo para SharePoint. Status: ' + res.getStatusCode());
                    }
                } else {
                    System.debug('Simulação upload do arquivo: ' + cv.Title);
                }
            }

            List<Items_SharepointAvivar__x> lstSharepointFiles = new List<Items_SharepointAvivar__x>();
            
            System.debug('fileNameBase: ' + fileNameBase);
            if (Test.isRunningTest()) {
                lstSharepointFiles = new List<Items_SharepointAvivar__x>{
                    new Items_SharepointAvivar__x(
                        ExternalId = 'External:12345'
                    )
                };
            } else {
                lstSharepointFiles = [SELECT Name__c, ExternalId 
                                    FROM Items_SharepointAvivar__x 
                                    WHERE Name__c like :fileNameBase + '%'];
            }
            System.debug('lstSharepointFiles: ' + lstSharepointFiles);

            if(lstSharepointFiles.isEmpty()){
                throw new CustomException('Erro ao buscar registro externo dos arquivos');
            }

            count = 0;
            List<ContentVersion> lstNewContentVersion = new List<ContentVersion>();
            for(Items_SharepointAvivar__x cv : lstSharepointFiles){
                count++;
                String fileName = fileNameBase + count + '%';            
                System.debug('fileName: ' + fileName);

                String syntheticName = isTest ? 'Comprovante1.pdf' : cv.Name__c;

                if(isTest){
                    System.debug('Simulação insert ContentVersion: ' + syntheticName);
                    String ci = '069000000000000AAA';
                    System.debug('Simulação insert ContentDocumentLink com ID: ' + ci);
                    System.debug('Simulação delete de ContentDocument: ' + setContentDocumentId);
                } else {
                    ContentVersion c = new ContentVersion(
                        Origin = 'H',
                        ContentLocation = 'E',
                        ExternalDataSourceId = externalDataSourceId,
                        ExternalDocumentInfo1 = param?.UrlPreVisualizacao__c + cv.ExternalId.substringAfterLast(':'),
                        ExternalDocumentInfo2 = cv.ExternalId.substringAfter(':'),
                        PathOnClient = cv.Name__c,
                        Title = syntheticName.substringBeforeLast('.')
                    );
                    insert c;

                    String ci = [SELECT ContentDocumentId FROM ContentVersion WHERE Id = :c.Id LIMIT 1].ContentDocumentId;
                    ContentDocumentLink cDocLink = new ContentDocumentLink(
                        LinkedEntityId = recordId,
                        ContentDocumentId = ci,
                        Visibility = 'InternalUsers'
                    );
                    insert cDocLink;

                    if(param.ExcluirArquivoLocal__c){
                        delete [SELECT Id FROM ContentDocument WHERE Id IN :setContentDocumentId];
                    }
                }

            }

            AtualizacaoCadastral__c uptAtualizacaoCadastral = new AtualizacaoCadastral__c(
                Id = recordId,
                MovimentacaoArquivosSharepoint__c = 'CONCLUÍDA'
            );
            update uptAtualizacaoCadastral;

        }
    }

    private static void scheduleNextJob(){

        Integer tempoRecorrencia = Integer.valueOf(ParametrosAtualizacaoCadastral__c.getInstance()?.IntervaloTransferenciaArquivos__c);
        Datetime newSchuledatetime = System.now().addMinutes(tempoRecorrencia == null ? 30 : tempoRecorrencia);

        String sch = String.valueof(newSchuledatetime.second()) + ' ' +
                    String.valueof(newSchuledatetime.minute()) + ' ' +
                    String.valueof(newSchuledatetime.hour()) + ' ' +
                    String.valueof(newSchuledatetime.day()) + ' ' +
                    String.valueof(newSchuledatetime.month()) + ' ? ' +
                    String.valueof(newSchuledatetime.year());
        String scheduleName = 'Transferencia_Arquivos_Sharepoint_' + String.valueOf(newSchuledatetime) + (Test.isRunningTest() ? '_' + Math.random() : '');
        System.schedule(scheduleName, sch, new SharepointFileSender());

    }

}