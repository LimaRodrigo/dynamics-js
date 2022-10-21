/**
 * 18/11/2020
 * HelperXrm Utils 2.0
 * 
 * ******************** ATENÇÃO ********************
 * bibliotéca só se aplica a navegadores modernos, IE não tem compatibilidade, em caso assim utilizar o PU antigo.
 * 
 * ******************** REGRAS DE NOMENCLATURA ******************** 
 * 
 * [Funções]
 * > Padrão Pascal Case
 * > Que retornem algo começar com a palavra "Retornar"
 * > Boleanas que retonam algo começar com a palavra "Verificar"
 * > Que inserem dados começar com a palavra "Inserir"
 * > Que definem comportamento começar com a palavra "Definir"
 * > Funções assincronas ou que tenham Promisse colocar a palavra [Async] no final. Ex.: RetornarRegistrosAsync 
 * > Funçoes que facilitam alguma outra função, mas não é utilizada extenamente, criar no Objeto Helper
 * > Atentar-se para utilização do this, para evitar erros
 * 
 * [Propriedades]
 * > Padrão Camel Case
 * 
 * [Parâmetros]
 * > Padrão Camel Case
 * 
 * [variaveis]
 * > Padrão Camel Case
 * > variaveis locais declarar como let.
 * > Não criar varieveis globais, gerar um proriedade no utils.
 */

if (typeof (HelperXrm) === "undefined") { HelperXrm = {}; }
/**
 * Bibliotéca para auxilio na utilização do SDK JS Dynamics CRM 365
 */
HelperXrm.Utils = {

    formContext: {},
    globalContext: Xrm.Utility.getGlobalContext(),

    /**
     * Atualiza os dados da tela
     * @param {boolean} salvar - Define o salvamento antes do refresh ou não
     * @param {object} successCallback - Função de callback em caso de sucesso
     * @param {object} errorCallback - Função de callback em caso de erro
     */
    AtualizarCamposDaTela: function (salvar, successCallback, errorCallback) {
        this.formContext.data.refresh(salvar).then(successCallback, errorCallback);
    },

    /**
     * Retorna se o usuário está navegando no UCI (HUB)
     * @returns {boolean}
     */
    VerificarEUCI: function () {
        return Xrm.Internal.isUci();
    },

    /**
     * Salvar os dados do form
     * @param {object} successCallback - Função de callback em caso de sucesso
     * @param {object} errorCallback - Função de callback em caso de erro
     */
    SalvarForm: function (successCallback, errorCallback) {
        this.formContext.data.save(null).then(successCallback, errorCallback);
    },

    /**
     * Insere ou limpa notificações do form.
     * @param {string} uniqueId - Identificador usado para limpar a notificação
     * @param {string} level - ERROR / WARNING / INFO
     * @param {string} message - Texto da mensagem 
     */
    ExibirNotificacaoForm: function (uniqueId, level = null, message = null) {
        message ? this.formContext.ui.setFormNotification(message, level, uniqueId) : this.formContext.ui.clearFormNotification(uniqueId);
    },

    /**
     * Insere ou limpa notificações do campo
     * @param {string} uniqueId - Identificador usado para limpar a notificação
     * @param {string} campo - Nome lógico do campo
     * @param {string} message - Texto da mensagem 
     */
    ExibirNotificacoesDoCampo: function (uniqueId, campo, message = null) {

        message ? this.formContext.getControl(campo).setNotification(message, uniqueId) : this.formContext.getControl(campo).clearNotification(uniqueId);
    },
    /**
     * Retorna o uniquename da organização
     * @returns {string}
     */
    RetornarNomeDaOrganizacao: function () {
        return this.globalContext.organizationSettings.uniqueName;
    },

    /**
     * Retorna o Id da organização
     * @returns {string}
     */
    RetornarIdDaOrganizacao: function () {
        return this.globalContext.organizationSettings.organizationId;
    },
    /**
     * Retorna o codigo da linguagem da organização. Ex: pt-br = 1046 
     * @returns {number}
     */
    RetornarCondigoLinguagemDaOrganizacao: function () {
        return this.globalContext.organizationSettings.languageId;
    },
    /**
     * Retorna id do usuário logado
     * @returns {string}
     */
    RetornarIdDoUsuario: function () {
        return this.globalContext.userSettings.userId;
    },
    /**
     * Retorna o codigo da linguagem do usuário logado. Ex: pt-br = 1046 
     * @returns {number}
     */
    RetornarCodigoLinguagemDoUsuario: function () {
        return this.globalContext.userSettings.languageId;
    },
    /**
     * Retorna lista ids dos Diretos de acesso do usuário logado
     * @returns {Array}
     */
    RetornarIdsDireitoAcessoDoUsuario: function () {
        return this.globalContext.userSettings.securityRoles;
    },
    /**
     * Retona o id do registro
     * @returns {string}
     */
    RetornarIdDoRegistro: function () {
        return this.formContext.data.entity.getId();
    },
    /**
     * Retorna o controle do iframe
     * @param {string} nomeIframe - Nome Lógico do Iframe no form
     * @returns {object}
     */
    RetornarControleIframe: function (nomeIframe) {
        return this.formContext.ui.controls.get(nomeIframe);
    },
    /**
     * Insere filtro customizado nos campos do tipo lookUp
     * @param {string} campo - Nome lógico do campo lookup que será filtrado
     * @param {object} funcao - Função que constrói o filtro
     */
    InserirFiltroCustomizado: function (campo, funcao) {
        this.formContext.getControl(campo).addPreSearch(funcao);
    },

    /**
     * Retornar Verdadeiro caso o formulário seja do tipo Atualização
     */
    VerificarTipoDoFormAtualizacao: function () {
        this.formContext.getFormType() == 2;
    },
    /**
     * Retornar Verdadeiro caso o formulário seja do tipo Criaçao
     */
    VerificarTipoDoFormCriacao: function () {
        this.formContext.getFormType() == 1;
    },
    /**
    * Retornar Verdadeiro caso o formulário seja do tipo Somente Leitura
    */
    VerificarTipoDoFormSomenteLeitura: function () {
        this.formContext.getFormType() == 3;
    },
    /**
    * Retornar Verdadeiro caso o formulário seja do tipo Desabilitado
    */
    VerificarTipoDoFormDesabilitado: function () {
        this.formContext.getFormType() == 4;
    },
    /**
     * Retorna Vedadeiro caso haja alterações não salvas no campo.
     * @param {string} campo - nome lógico do campo
     * @returns {boolean}
     */
    VerificarAlteracaoNaoSalvaNoCampo: function (campo) {
        return this.formContext.getAttribute(campo).getIsDirty();
    },
    /**
     * Retorna true caso tenha algum campo que não foi salvo.
     * @returns {boolean}
     */
    VerificarAlteracaoTodosCamposNaoSalvos: function () {
        let camposPendente = false;
        let oppAttributes = this.formContext.data.entity.attributes.get();
        for (let i in oppAttributes) {
            if (oppAttributes[i].getIsDirty())
                camposPendente = true;
        }
        return camposPendente;
    },
    /**
     * Retorna o valor do campo
     * @param {string} campo - Nome lógico do campo
     * @returns {any}
     */
    RetornarValor: function (campo) {
        try {
            return this.formContext.getAttribute(campo).getValue();
        }
        catch (e) {
            console.log(e);
            return null;
        }
    },
    /**
     * Retorna o valor do campo
     * @param {string} campo - Nome lógico do campo
     * @param {object} valor - Valor a ser inserido
     */
    InserirValor: function (campo, valor) {
        try {
            this.formContext.getAttribute(campo).setValue(valor);
        }
        catch (e) {
            console.log(e);
            return null;
        }
    },
    /**
     * Define o foco no campo desejado.
     * @param {string} campo - Nome lógico do campo
     */
    DefinirFocoCampo: function (campo) {
        this.formContext.getControl(campo).setFocus();
    },
    /**
     * Desabilita ou Habilita campo de acordo com "desabilitar" true ou false
     * @param {string} campo - Nome lógico do campo
     * @param {boolean} desabilitar - true ou false define se habilita ou desabilita o campo
     */
    DesabilitarCampo: function (campo, desabilitar) {
        this.formContext.getControl(campo).setDisabled(desabilitar);
    },
    /**
     * Desabilita todos os campos com possibilidade de execeções.
     * @param {boolean} desabilitar true ou false define se habilita ou desabilita os campos
     * @param {Array} listaExcecoes - Array string com nome lógico dos campos que não deve ser afatados
     */
    DesabilitarTodosCampos: function (desabilitar, listaExcecoes = []) {
        let attributes = this.formContext.data.entity.attributes.get();
        for (i = 0; i < attributes.length; i++) {
            if (this.formContext.getControl(attributes[i].getName()) != null)
                if ((listaExcecoes.indexOf(attributes[i].getName())) == -1)
                    this.DesabilitarCampo(attributes[i].getName(), desabilitar);
        }
    },
    /**
     * Defini o comportamento do salvamento de dados de um campo
     * @param {string} campo - Nome lógico do campo
     * @param {string} modo - Possibilidade: [always] Dados sempre salvos, [never] Dados nuca serão salvos, [dirty] Dados salvos caso haja ateração (default)
     */
    DefinirModoSalvar: function (campo, modo) {
        this.formContext.getAttribute(campo).setSubmitMode(modo);
    },

    /**
     * Defini o nivel de requerimento, podendo ser obrigatorio ou apenas recomendado
     * @param {string} campo - Nome lógico do campo
     * @param {string} nivel - Nivel de recomendação possíveis:  none | required |recommended
     */
    DefinirNivelRecomendacao: function (campo, nivel) {
        this.formContext.getAttribute(campo).setRequiredLevel(nivel);
    },
    /**
     * Defini visibilidade do campo
     * @param {string} campo - Nome lógico do campo
     * @param {boolean} visivel - True ou false para exibir ou ocultar o campo 
     */
    DefinirVisibilidade: function (campo, visivel) {
        this.formContext.getControl(campo).setVisible(visivel);
    },
    /**
     * Devine a visbilidade de uma seção
     * @param {string} guia - Nome lógico da Guia
     * @param {string} secao - Nome lógico da Seção
     * @param {boolean} visivel - True ou false para exibir ou ocultar o campo
     */
    DefinirVisibiliadeSecao: function (guia, secao, visivel) {
        this.formContext.ui.tabs.get(guia).sections.get(secao).setVisible(visivel);
    },
    /**
     * Devine a visbilidade de uma guia
     * @param {string} guia - Nome lógico da Guia
     * @param {boolean} visivel - True ou false para exibir ou ocultar o campo
     */
    DefinirVisibiliadeGuia: function (guia, visivel) {
        this.formContext.ui.tabs.get(guia).setVisible(visivel);
    },
    /**
     * Defini um função no onchange de um campo
     * @param {string} campo - Nome lógico do campo
     * @param {object} funcao  - Sua função que será disparada no onchange
     */
    DefinirMetodoOnChange: function (campo, funcao) {
        this.formContext.getAttribute(campo).addOnChange(funcao);
    },
    /**
     * Retorna o nome do form atual
     * @returns {string}
     */
    RetornarNomeFormAtual: function () {
        return this.formContext.ui.formSelector.getCurrentItem().getLabel();
    },
    /**
     * Retorna o nome da entidade do formulário
     * @returns {string}
     */
    RetornarNomeDaEntidade: function () {
        return this.formContext.data.entity.getEntityName();
    },

    /**
     * Exibe um alerta modal na tela
     * @param {string} mensagem - Mensagem a ser exibida
     * @param {string} titulo - titulo do modal
     * @param {number} largura - tamanho lagura medido em pixels, caso nulo padrão 350
     * @param {number} altura - tamanho altura medido em pixels,  caso nulo padrão 550
     */
    ExibirAlertaSemConfirmacao: function (mensagem, titulo = null, largura = null, altura = null) {
        let alertStrings = {
            title: titulo == null ? "" : titulo,
            confirmButtonLabel: "OK",
            text: mensagem
        };
        let alertOptions = {
            height: altura == null ? 350 : altura,
            width: largura == null ? 550 : largura
        };
        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
            function success(result) {
                console.log("Sucesso! Alerta exibido!");
            },
            function (error) {
                concole.log(error.message);
            }
        );
    },
    /**
     * Exibe um alerta modal na tela com um confirmação possibilitando um ação após a execução.
     * @param {string} mensagem - Mensagem a ser exibida
     * @param {object} funcaoSucesso - função a ser excutada após a confirmação. obs.: a função não pode conter parametros
     * @param {string} titulo - titulo do modal
     * @param {string} subtitulo - Subtitulo do modal
     * @param {number} largura - tamanho lagura medido em pixels, caso nulo padrão 350
     * @param {number} altura - tamanho altura medido em pixels,  caso nulo padrão 550
     */
    ExibirAlertaComConfirmacao: function (mensagem, funcaoSucesso, titulo = null, subtitulo = null, largura = null, altura = null) {
        let confirmStrings = {
            cancelButtonLabel: "Cancelar",
            confirmButtonLabel: "Confirmar",
            text: mensagem,
            title: titulo == null ? "" : titulo,
            subtitle: subtitulo == null ? "" : subtitulo
        };

        let confirmOptions = {
            height: altura == null ? 350 : altura,
            width: largura == null ? 550 : largura
        };

        Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
            (success) => {
                if (success.confirmed)
                    funcaoSucesso();

                else
                    console.log("Usuário cancelou a ação");
            },
            (error) => {
                concole.log(error.message);
            });
    },
    /**
     * Exibe ou oculta o loading
     * @param {boolean} exibir - true exibe, false oculta
     * @param {string} mensagem - mensagem a ser exibida 
     */
    ExibirOcultarLoading: function (exibir, mensagem = null) {
        if (exibir)
            Xrm.Utility.showProgressIndicator(mensagem);
        else
            Xrm.Utility.closeProgressIndicator();
    },
    /**
     * Exibe o loading e oculta de acordo com tempo informado
     * @param {numbe} tempo - Tempo em milisegundos 
     * @param {string} mensagem - mensagem a ser exibida
     */
    ExibirLoadingComTimer: function (tempo, mensagem = null) {
        setTimeout(() => { Xrm.Utility.showProgressIndicator(mensagem); }, tempo);
        Xrm.Utility.closeProgressIndicator();
    },
    /**
     * Verifica se existe algum campo obrigatório não preenchido
     * @returns {boolean} Retorna true caso estejam todos preenchidos.
     */
    VerificarPreenchimentoCamposObrigatorios: function () {
        let populated = true;
        this.formContext.getAttribute(function (attribute) {
            if (attribute.getRequiredLevel() === "required") {
                if (attribute.getValue() === null) {
                    populated = false;
                }
            }
        });
        return populated;
    },

    /**
     * Verifica se trexo contem no texto completo
     * @param {string} textoCompleto - Texto campo a qual ira verificar se palavra contem
     * @param {string} valor - valor a ser procurado
     * @returns {boolean}
     */
    VerificaTextoContem: function (textoCompleto, valor) {
        let str = textoCompleto.toLowerCase();
        return str.indexOf(valor.toLowerCase()) >= 0 ? true : false;
    },
    /**
     * Compara duas datas, retorna TRUE se a segunda data for MAIOR do que a primeira
     * @param {Date} primeiraData 
     * @param {Date} segundaData
     * @returns {boolean} 
     */
    CompararDatas: function (primeiraData, segundaData) {
        let result = false;

        let dataPrimaria = { dia: 0, mes: 0, ano: 0, total: 0 };
        let dataSecundaria = { dia: 0, mes: 0, ano: 0, total: 0 };

        dataPrimaria.dia = parseInt(primeiraData.getDate().toString());
        dataPrimaria.mes = parseInt(primeiraData.getMonth().toString());
        dataPrimaria.ano = parseInt(primeiraData.getYear().toString());
        dataPrimaria.total = dataPrimaria.dia * 10 + Math.pow(dataPrimaria.mes * 100, 2) + Math.pow(dataPrimaria.ano * 100, 2);

        dataSecundaria.dia = parseInt(segundaData.getDate().toString());
        dataSecundaria.mes = parseInt(segundaData.getMonth().toString());
        dataSecundaria.ano = parseInt(segundaData.getYear().toString());
        dataSecundaria.total = dataSecundaria.dia * 10 + Math.pow(dataSecundaria.mes * 100, 2) + Math.pow(dataSecundaria.ano * 100, 2);

        if (dataSecundaria.total > dataPrimaria.total)
            result = true;

        return result;
    },
    /**
     * Preenche a string à esquerda com o caractere fornecido, até que ela atinja o tamanho especificado.
     * @param {string} originalstr - Texto orginal
     * @param {number} strSize - tamanho maximo
     * @param {string} addChar - Texto a ser adcionado
     * @returns {string}
     */
    PreencherAEsquerda: function (originalstr, strSize, addChar) {
        while (originalstr.length < strSize)
            originalstr = addChar + originalstr;
        return originalstr;
    },
    /**
     * retorna somente numeros 
     * @param {string} texto 
     * @returns {string}
     */
    RetornarSomenteNumeros: function (texto) {
        return texto.replace(/\D/g, '');
    },
    /**
     * Verifica se o valor é numero 
     * @param {any} valor
     * @returns {boolean}
     */
    VerificarENumero: function (valor) {
        return !isNaN(parseFloat(valor)) && isFinite(valor);
    },
    /**
     * Modifica todos os campos texto pra maiúsculo
     */
    ModificarTodosCampoMaiusculos: function () {
        let controls = HelperXrm.Utils.formContext.ui.controls.get();
        for (i = 0; i < controls.length; i++) {
            try {
                let control = controls[i];
                if (control.getControlType() == "standard") {
                    if (typeof (control.getAttribute) !== "function") continue;
                    let attribute = control.getAttribute();
                    if (attribute && attribute.getAttributeType() &&
                        (attribute.getAttributeType() == "string" || attribute.getAttributeType() == "memo")) {
                        let valor = attribute.getValue();
                        if (valor)
                            attribute.setValue(valor.toUpperCase());
                    }
                }
            } catch (error) {
                console.log(error);
                continue;
            }

        }
    },
    /**
     * Executar uma ação, exemplo ação de um botão. 
     * O erro será tratado internamente com possibilidade de mensagem personalizada.
     * O sucesso ira retornar o objeto com os paramentros de saida configurados no ação.
     * Utilize o Await ao chamar a função para aguardar o termino da função.
     * @param {string} id - id do registros, ex.: {CFF5AA7F-B426-EB11-BBF3-000D3A887C31}
     * @param {string} nomeEntidade - Nome lógico da entidade, ex.: account
     * @param {string} nomeAcao - Nome lógico da ação ex.: ptr_Acao
     * @param {string} msgErro - Mensagem exibida caso gere erro no processo, caso não seja definida a mensagem de erro será o proprio retorno da Action, a exemplo a exception do plugin  
     * @returns {object}
     */
    ExecutarAcaoAsync: async function (id, nomeEntidade, nomeAcao, msgErro = null) {
        let retorno;
        let target = {};
        target.entityType = nomeEntidade;
        target.id = id.replace('{', '').replace('}', '');
        let req = {};
        reqreq.entity = target;
        req.getMetadata = function () {
            return {
                boundParameter: "entity",
                parameterTypes: {
                    "entity": {
                        "typeName": "mscrm." + nomeEntidade,
                        "structuralProperty": 5
                    }
                },
                operationType: 0,
                operationName: nomeAcao
            };
        };


        try {
            await HelperXrm.Helper.Savar();
            HelperXrm.Utils.ExibirOcultarLoading(true, "Enviando");
            retorno = await HelperXrm.Helper.ExecutarRequest(req);
            HelperXrm.Utils.ExibirOcultarLoading(false);

        } catch (error) {
            console.log(error.message)
            if (msgErro !== null && msgErro !== undefined)
                HelperXrm.Utils.ExibirAlertaSemConfirmacao(msgErro, "Erro", 700, 400);
            else if (error.message !== null && error.message !== undefined)
                HelperXrm.Utils.ExibirAlertaSemConfirmacao(error.message, "Erro", 700, 400);
            else
                HelperXrm.Utils.ExibirAlertaSemConfirmacao("Erro no processo.", "Erro", 700, 400);
            HelperXrm.Utils.ExibirOcultarLoading(false);

        }
        return retorno;
    },

    /**
     * Retorna um array de objetos de acordo com a busca, na chamada é necessário chamar dentro de um try-catch e utilizar await na frente para aguardar o resultado. 
     * Mais detalhes no Link: 
     * https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/xrm-webapi/retrievemultiplerecords
     * @param {string} nomeEntidade - Nome Lógico da Entidade 
     * @param {string} optionOdata - Passa estruta odata completa com select, filter, top, etc ex.:?$select=name&$filter=campo eq 'teste' &$top=2
     * @returns {object}
     */
    RetornarRegistrosAsync: function (nomeEntidade, optionOdata) {
        return new Promise((resolve, reject) => {
            Xrm.WebApi.retrieveMultipleRecords(nomeEntidade, optionOdata).then(
                (result) => { resolve(result.entities); },
                (e) => { reject(e); }
            );
        });
    },
    /**
     * Retorna apenas o registros de acordo com o Id, na chamada é necessário chamar dentro de um try-catch e utilizar await na frente para aguardar o resultado. 
     * Mais detalhes no Link: 
     * https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/xrm-webapi/retrieverecord
     * @param {string} nomeEntidade - Nome lógico da entidade
     * @param {string} id - Id do registro a ser recuperado ex.: CFF5AA7F-B426-EB11-BBF3-000D3A887C31
     * @param {string} optionOdata -  Passa estruta odata select e expand(opcional) ex.:  "?$select=name&$expand=primarycontactid($select=contactid,fullname)"
     * @returns {object}
     */
    RetornarRegistroAsync: function (nomeEntidade, id, optionOdata) {
        return new Promise((resolve, reject) => {
            Xrm.WebApi.retrieveRecord(nomeEntidade, id, optionOdata).then(
                (result) => { resolve(result); },
                (e) => { reject(e); }
            );
        });
    },
    /**
     * Cria um registro com possiblidade de relacionar registro e/ou criar registros relcionados, na chamada é necessário chamar dentro de um try-catch e utilizar await na frente para aguardar o resultado. 
     * Mais detalhes no Link: 
     * https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/xrm-webapi/createrecord
     * @param {string} nomeEntidade - Nome lógico da entidade
     * @param {object} objeto - Objeto a ser criado
     * @returns {object}
     */
    CriarRegistroAsync: function (nomeEntidade, objeto) {
        return new Promise((resolve, reject) => {
            Xrm.WebApi.createRecord(nomeEntidade, objeto).then(
                (result) => { resolve(result); },
                (e) => { reject(e); }
            );
        });
    },
    /**
     * Atualiza um Resgitro, na chamada é necessário chamar dentro de um try-catch e utilizar await na frente para aguardar o resultado. 
     * Mais detalhes no Link: 
     * https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/xrm-webapi/updaterecord
     * @param {string} nomeEntidade - Nome lógico da entidade
     * @param {string} id Id do registro a ser atualizado ex.: CFF5AA7F-B426-EB11-BBF3-000D3A887C31
     * @param {object} objeto - Objeto a ser atualizado
     * @returns {object}
     */
    AtualizarRegistroAsync: function (nomeEntidade, id, objeto) {
        return new Promise((resolve, reject) => {
            Xrm.WebApi.updateRecord(nomeEntidade, id, objeto).then(
                (result) => { resolve(result); },
                (e) => { reject(e); }
            );
        });
    },
    /**
     * Deleta um registro
     * @param {string} nomeEntidade - Nome lógico da entidade
     * @param {string} id Id do registro a ser atualizado ex.: CFF5AA7F-B426-EB11-BBF3-000D3A887C31
     */
    DeletarRegistroAsync: function (nomeEntidade, id) {
        return new Promise((resolve, reject) => {
            Xrm.WebApi.deleteRecord(nomeEntidade, id).then(
                (result) => { resolve(result); },
                (e) => { reject(e); }
            );
        });
    },
    /**
     * Gera um objeto de requisição de criação
     * @param {string} nomeEntidade - - Nome lógico da entidade
     * @param {object} objeto - objeto do registro a ser criado
     * @returns {object}
     */
    RetornarRequisicaoCriacao: function (nomeEntidade, objeto) {
        let CreateRequest = {};
        CreateRequest.etn = nomeEntidade;
        CreateRequest.payload = objeto;
        CreateRequest.getMetadata = () => {
            return {
                boundParameter: null,
                parameterTypes: {},
                operationType: 2,
                operationName: "Create",
            };
        }
        return CreateRequest;
    },
    /**
     * Gera um objeto de requisição de atualização
     * @param {string} nomeEntidade - - Nome lógico da entidade
     * @param {string} id - id do registro sem "{ }"
     * @param {object} objeto - objeto do registro a ser criado
     * @returns {object}
     */
    RetornarRequisicaoAtualizacao: function (nomeEntidade, id, objeto) {
        let UpdateRequest = {};
        UpdateRequest.etn = nomeEntidade;
        UpdateRequest.id = id;
        UpdateRequest.payload = objeto;
        UpdateRequest.getMetadata = () => {
            return {
                boundParameter: null,
                parameterTypes: {},
                operationType: 2,
                operationName: "Update",
            };
        }
        return UpdateRequest;
    },
    /**
     * Gera um objeto de requisição de Deleção
     * @param {string} nomeEntidade - - Nome lógico da entidade
     * @param {string} id - id do registro sem "{ }"
     * @returns {object}
     */
    RetornarRequisicaoDelecao: function (nomeEntidade, id) {
        let DeleteRequest = {};
        DeleteRequest.entityReference = {
            entityType: nomeEntidade,
            id: id
        };
        DeleteRequest.getMetadata = () => {
            return {
                boundParameter: null,
                parameterTypes: {},
                operationType: 2,
                operationName: "Delete",
            };
        }
        return DeleteRequest;
    },
    /**
     * retorna o request montado de acordor com a operação desejada
     * https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/xrm-webapi/online/execute
     * @param {string} nomeEntidade - Nome lógico da entidade
     * @param {object} listaObjetos - Lista de objetos 
     * @param {string} nomeOperacao - Possíveis: Create | Update | Delete . 
     * @param {boolean} eAtividade - enviar true caso registros sejam atividades, por padrão é false
     * @returns {Array object}
     */
    RetornarListaDeRequisicaoes: function (nomeEntidade, listaObjetos, nomeOperacao, eAtividade = false) {
        let arrayRequests = [];
        let request = {};
        for (let i = 0; i < listaObjetos.length; i++) {
            const entity = listaObjetos[i];
            request = {};

            if (nomeOperacao == "Create" || nomeOperacao == "Update") {
                request.etn = nomeEntidade;
                request.payload = entity;
            }
            if (nomeOperacao == "Update")
                request.id = eAtividade ? entity.activityid : entity[nomeEntidade + "id"];
            if (nomeOperacao == "Delete")
                request.entityReference = { entityType: nomeEntidade, id: eAtividade ? entity.activityid : entity[nomeEntidade + "id"] };

            request.getMetadata = () => {
                return {
                    boundParameter: null,
                    parameterTypes: {},
                    operationType: 2,
                    operationName: nomeOperacao,
                };
            };
            arrayRequests.push(request);
        }
        return arrayRequests;
    },
    /**
     * Executa mutipleRequest Retornando true caso sucesso.
     * @param {object} listaRequisicoes - Lista de requisições, deverá ser gerado na Function [RetornarListaDeRequisicaoes]
     * @returns {boolean} Retorna true caso sucesso.
     */
    ExecutarMultiplasRequisicoesAsync: async function (listaRequisicoes) {
        return await HelperXrm.Helper.ExecutarMutiplosRequest(listaRequisicoes);
    },
    /**
    * Executa request Retornando true caso sucesso.
    * @param {object} listaRequisicoes - A requisiçao, deverá ser gerado nas funções que geram request de Create, Update e Delete
    * @returns {boolean} Retorna true caso sucesso.
    */
    ExecutarRequisicaoAsync: async function (requisicao) {
        return await HelperXrm.Helper.ExecutarRequest(requisicao);
    },
    /**
     * Retorna dados via api sincronamente. 
     * Utilizar essa função somente em casos de Filtros de Lookup, devido funções async não funcionarem. 
     * Em outros casos utilizar função [RetornarRegistros]
     * @param {string} url - url completa ex.:https://trialdynamicsnovembro.crm2.dynamics.com/api/data/v9.1/accounts
     * @returns {object}
     */
    RetornarRegistrosViaApi: function (url) {
        let dados;
        let erro = "";
        jQuery.ajax({
            type: "GET",
            contentType: "application/json; charset=utf-8",
            datatype: "json",
            url: url,
            async: false,
            beforeSend: function (XMLHttpRequest) {
                XMLHttpRequest.setRequestHeader("Accept", "application/json");
                XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=*");
            },
            success: function (data, textStatus, XmlHttpRequest) {
                if (data) dados = data;
            },
            error: function (XmlHttpRequest, textStatus, errorThrown) {
                if (XmlHttpRequest.responseJSON && XmlHttpRequest.responseJSON.error && XmlHttpRequest.responseJSON.error.message) {
                    console.log(XmlHttpRequest.responseJSON.error.message);
                    erro = XmlHttpRequest.responseJSON.error.message;
                }
            }
        });
        if (erro != "")
            throw erro;
        return dados;
    },
    /**
     * Retorna url do ambeinte ex.:https://contoso.crm2.dynamics.com
     * @returns {string}
     */
    RetornarUrlAmbiente: function () {
        this.globalContext.getClientUrl();
    },
    /**
     * Retorna o valor do parametro na url
     * @param {string} nomeParametro - Nome do paramentro a ser buscado
     * @returns {string}
     */
    RetornarParametroDeUrl: function (nomeParametro) {
        let valores = location.search.substr(1).split("&");
        for (let i in valores)
            valores[i] = valores[i].replace(/\+/g, " ").split("=");
        for (let j in valores) {
            if (valores[j][0].toLowerCase() === "data") {
                let valoresDecodificados = decodeURIComponent(valores[j][1]).split("&");
                for (let k in valoresDecodificados) {
                    let valor = valoresDecodificados[k].replace(/\+/g, " ").split("=");

                    if (valor[0] === nomeParametro)
                        return valor[1];
                }
            }
        }
        return null;
    },
    /**
     * Grupo de funções para campos lookup
     */
    Lookup: {
        /**
         * Monta o objeto pra que possa gravar um valor lookup, na função Inserir valor
         * @param {string} id - Id do registro relacionado
         * @param {string} entityType - nome lógico da entidade relacionada
         * @param {string} name - Texto do campo name da entidade relacionada
         */
        MontarObjeto: function (id, entityType, name) {
            try {
                let lookupData = new Array();
                let lookupItem = new Object();
                lookupItem.id = id;
                lookupItem.entityType = entityType;
                lookupItem.name = name;
                lookupData[0] = lookupItem;
                return lookupData;
            } catch (e) {
                console.log(e);
                return null;
            }
        },

        /**
         * Retorna o id de um campo lookup
         * @param {string} campo  - Nome lógico do campo
         */
        RetornarGuid: function (campo) {
            try {
                return HelperXrm.Utils.formContext.getAttribute(campo).getValue()[0].id;
            } catch (e) {
                console.log(e);
                return null;
            }
        },
        /**
         * Retorna o Nome (texto) de um campo lookup
         * @param {string} campo - Nome lógico do campo
         */
        RetornarNome: function (campo) {
            try {
                return HelperXrm.Utils.formContext.getAttribute(campo).getValue()[0].name;
            } catch (e) {
                console.log(e);
                return null;
            }
        },
        /**
         * Insere um função no pre filtro do lookup (addPreSearch).
         * Essa função deve ser utilizada somente no onload, a função não pode ser async 
         * ou ter algum promisse exemplo utilização de Xrm.WebApi
         * @param {string} campo - Nome lógico do campo
         * @param {object} funcao - função referencia
         */
        InserirFuncaoPreFiltro: function (campo, funcao) {
            HelperXrm.Utils.formContext.getControl(campo).addPreSearch(funcao);
        },
        /**
         * Adciona um filtro customizado em um lookup, lembrando que a função que chama está função 
         * terá que ser referenciada no onload com a função HelperXrm.Utils.InserirFuncaoPreFiltroLookup
         * @param {string} campo - Nome lógico do campo
         * @param {string} fetchXml  - exemplo <filter type='and'><condition attribute='accountid' operator='in'><value>c5f5aa7f-b426-eb11-bbf3-000d3a887c31</value> </condition></filter>
         * @param {string} nomeEntidade 
         */
        InserirFiltroCustomizado: function (campo, fetchXml, nomeEntidade) {
            HelperXrm.Utils.formContext.getControl(campo).addCustomFilter(fetchXml, nomeEntidade);
        },

    },
    /**
     * Grupo de funções somente para campos OptionSet
     */
    Optionset: {
        /**
         * Retorna o valor id que estão selecionado
         * @param {string} campo - Nome lógico do campo
         */
        RetornarValorSelecionado: function (campo) {
            try {
                return HelperXrm.Utils.formContext.getAttribute(campo).getSelectedOption().value;
            } catch (e) {
                console.log(e);
                return null;
            }
        },
        /**
        * Retorna uma lista de ids selecionados
        * @param {string} campo - Nome lógico do campo
        */
        RetornarValorSelecionadoMultiplo: function (campo) {
            try {
                let Values = new Array();
                let opValue = HelperXrm.Utils.formContext.getAttribute(campo).getSelectedOption();
                opValue.forEach(element => { Values.push(element.value); });
                return Values;
            } catch (e) {
                console.log(e);
                return null;
            }
        },
        /**
         * Retorna lista de objetos com todas opções 
         * @param {string} campo - Nome lógico do campo
         */
        RetornarTodasOpcoes: function (campo) {
            return HelperXrm.Utils.formContext.getAttribute(campo).getOptions();
        },
        /**
         * Retorna o valor label selecionado
         * @param {string} campo - Nome lógico do campo
         */
        RetornarTextoSelecionado: function (campo) {
            try {
                return HelperXrm.Utils.formContext.getAttribute(campo).getSelectedOption().text;
            } catch (e) {
                console.log(e);
                return null;
            }
        },
        /**
         * Retorna uma lista de nomes selecionados
         * @param {string} campo - Nome lógico do campo
         */
        RetornarTextoSelecionadoMultiplo: function (campo) {
            try {
                let Values = new Array();
                let opValue = HelperXrm.Utils.formContext.getAttribute(campo).getSelectedOption();
                opValue.forEach(element => { Values.push(element.text); });
                return Values;
            } catch (e) {
                console.log(e);
                return null;
            }
        },
        /**
         * Insere opção em um conjunto de oções, mas opção deve existir na base
         * @param {string} campo - Nome lógico do campo
         * @param {object} opcao - Objeto do tipo option exemplo 'let opcao = new Option(); opcao.value = 956840000; opcao.text = "Maçã";'
         * @param {number} index - Define a posição onde a opção será inserida
         */
        InserirOpcao: function (campo, opcao, index = 0) {
            HelperXrm.Utils.formContext.getControl(campo).addOption(opcao, index);
        },
        /**
         * Remove todas as opções do conjunto de opções
         * @param {string} campo - Nome lógico do campo
         */
        RemoverTodasOpcoes: function (campo) {
            HelperXrm.Utils.formContext.getControl(campo).clearOptions();
        },
        /**
         * Remove somente a opção passada
         * @param {string} campo - Nome lógico do campo
         * @param {number} valor - Valor inteiro da opção ex.:956840000
         */
        RemoverItemConjuntOpcao: function (campo, valor) {
            HelperXrm.Utils.formContext.getControl(campo).removeOption(valor);
        },

    },

    /**
     * Grupo de funções relacionados a grids do formulário
     */
    Grid: {
        /**
         * Atualiza (Refresh) o grind no formulário
         * @param {string} nomeGrid - Nome lógico do grid no formulário
         */
        AtualizarGridItens: function (nomeGrid) {
            HelperXrm.Utils.formContext.ui.controls.get(nomeGrid).refresh();
        },
        /**
         * Retorna uma listas com os registros do grid 
         * @param {string} NomeGrid - Nome lógico do grid no formulário
         * @returns array ex.: [{id:35F6AA7F-B426-EB11-BBF3-000D3A887C31, attributes: [{key:"name", value: "teste"}]}]
         */
        RetornarItensDoGrid: function (NomeGrid) {
            try {
                let retorno = new Array();
                let grid = HelperXrm.Utils.formContext.getControl(NomeGrid).getGrid();

                if (grid.getTotalRecordCount() == 0) { return retorno; }

                let rows = grid.getRows().getAll();

                for (let i = 0; i < rows.length; i++) {

                    let coll = rows[i].getData().entity.attributes.get();
                    let linha = {};
                    let attributes = [];
                    for (let j = 0; j < coll.length; j++) {
                        attributes.push({ key: coll[j]._attributeName, value: coll[j].getValue() });
                    }
                    linha["id"] = rows[i].getData().entity.getId();
                    linha["attributes"] = attributes;

                    retorno.push(linha);
                }
                return retorno;
            } catch (e) {
                console.log(e);
                return null;
            }
        },
        /**
         * Retorna o total de itens no grid
         * @param {string} NomeGrid - Nome lógico do grid no formulário
         */
        RetornarTotalItensNoGrid: function (NomeGrid) {
            try {
                let grid = HelperXrm.Utils.formContext.getControl(NomeGrid).getGrid();
                return grid.getTotalRecordCount();
            } catch (e) {
                console.log(e);
                return null;
            }
        },
        /**
         * Defini um função no onload do grid desejado.
         * @param {string} NomeGrid - Nome lógico do grid no formulário
         * @param {object} funcao - Função a ser executada
         */
        DefinirFuncaoNoLoadDoGrid: function (NomeGrid, funcao) {
            HelperXrm.Utils.formContext.getControl(NomeGrid).addOnLoad(funcao);
        },
    },

    /**
     * Grupo de funções de Validação
     */
    Validadores: {
        /**
         * Valida o cpf
         * @param {string} cpf - texto do cpf, ex.: 000.000.000-00 ou 00000000000 
         */
        ValidarCPF: function (cpf) {
            cpf = cpf.replace(/\D/g, "");
            let numeros;
            let digitos;
            let soma;
            let i;
            let resultado;
            let digitos_iguais;
            digitos_iguais = 1;

            if (cpf.length != 11)
                return false;

            for (i = 0; i < cpf.length - 1; i++) {
                if (cpf.charAt(i) != cpf.charAt(i + 1)) {
                    digitos_iguais = 0;
                    break;
                }
            }

            if (!digitos_iguais) {
                numeros = cpf.substring(0, 9);
                digitos = cpf.substring(9);
                soma = 0;

                for (i = 10; i > 1; i--)
                    soma += numeros.charAt(10 - i) * i;

                resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;
                if (resultado != digitos.charAt(0))
                    return false;

                numeros = cpf.substring(0, 10);
                soma = 0;
                for (i = 11; i > 1; i--)
                    soma += numeros.charAt(11 - i) * i;
                resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;

                if (resultado != digitos.charAt(1))
                    return false;

                return true;
            }
            else {
                return false;
            }
        },
        /**
         * Valida o CNPJ
         * @param {string} cnpj - texto do cpf, ex.: 00.000.000/0000-0 ou 00000000000000 
         */
        ValidarCNPJ: function (cnpj) {
            cnpj = cnpj.replace(/\D/g, "");
            let numeros;
            let digitos;
            let soma;
            let i;
            let resultado;
            let pos;
            let tamanho;
            let digitos_iguais;
            digitos_iguais = 1;

            if (cnpj.length < 14 && cnpj.length < 15)
                return false;

            for (i = 0; i < cnpj.length - 1; i++) {
                if (cnpj.charAt(i) != cnpj.charAt(i + 1)) {
                    digitos_iguais = 0;
                    break;
                }
            }

            if (!digitos_iguais) {
                tamanho = cnpj.length - 2;
                numeros = cnpj.substring(0, tamanho);
                digitos = cnpj.substring(tamanho);
                soma = 0;
                pos = tamanho - 7;

                for (i = tamanho; i >= 1; i--) {
                    soma += numeros.charAt(tamanho - i) * pos--;
                    if (pos < 2)
                        pos = 9;
                }

                resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;
                if (resultado != digitos.charAt(0))
                    return false;

                tamanho = tamanho + 1;
                numeros = cnpj.substring(0, tamanho);
                soma = 0;
                pos = tamanho - 7;
                for (i = tamanho; i >= 1; i--) {
                    soma += numeros.charAt(tamanho - i) * pos--;
                    if (pos < 2)
                        pos = 9;
                }

                resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;
                if (resultado != digitos.charAt(1))
                    return false;

                return true;
            }
            else {
                return false;
            }
        },
        /**
         * Valida se o texto é um email
         * @param {string} email - texto do email 
         */
        ValidarEmail: function (email) {
            if (email == null)
                return;
            let conteudo = /^[\w-]+(\.[\w-]+)*@(([A-Za-z\d][A-Za-z\d-]{0,61}[A-Za-z\d]\.)+[A-Za-z]{2,6}|\[\d{1,3}(\.\d{1,3}){3}\])$/;
            return conteudo.test(email);
        },

    },
    /**
     * Grupo de função para aplicações de mascaras
     */
    Mascaras: {

        /**
         * Insere mascara para CPF ou CNPJ
         * @param {string} valor - numero cpf ou cnpj com ou sem pontuação ex.: CPF [000.000.000-00 ou 0000000000} 
         */
        InserirMascaraCPFouCNPJ: function (valor) {

            valor = valor.replace(/\D/g, ""); //Remove tudo o que não é dígito
            if (valor.length === 11) { //CPF
                valor = valor.replace(/(\d{3})(\d)/, "$1.$2");//Coloca um ponto entre o terceiro e o quarto dígitos
                valor = valor.replace(/(\d{3})(\d)/, "$1.$2");//Coloca um ponto entre o terceiro e o quarto dígitos
                valor = valor.replace(/(\d{3})(\d{1,2})$/, "$1-$2");  //Coloca um hífen entre o terceiro e o quarto dígitos
            }
            else if (valor.length === 14) { //CNPJ
                valor = valor.replace(/^(\d{2})(\d)/, "$1.$2"); //Coloca ponto entre o segundo e o terceiro dígitos
                valor = valor.replace(/^(\d{2})\.(\d{3})(\d)/, "$1.$2.$3"); //Coloca ponto entre o quinto e o sexto dígitos
                valor = valor.replace(/\.(\d{3})(\d)/, ".$1/$2"); //Coloca uma barra entre o oitavo e o nono dígitos
                valor = valor.replace(/(\d{4})(\d)/, "$1-$2");//Coloca um hífen depois do bloco de quatro dígitos
            }
            return valor;
        },

    },
};

/**
 * Grupo de funções facilitadoras para utilizalização interna do PU
 */
HelperXrm.Helper = {

    Savar: function () {
        return new Promise((resolve, reject) => {
            HelperXrm.Utils.formContext.data.save(null)
                .then((result) => { resolve(result); }, (e) => { reject(e); });
        });
    },
    ExecutarRequest: function (request) {
        return new Promise((resolve, reject) => {
            Xrm.WebApi.online.execute(request).then(
                (result) => {
                    resolve(result.ok);
                    // if (result.ok)
                    //     result.json().then((response) => { resolve(response); });
                },
                (error) => { reject(error); }
            );
        });
    },
    ExecutarMutiplosRequest: function (arrayRequest) {
        return new Promise((resolve, reject) => {
            Xrm.WebApi.online.executeMultiple(arrayRequest).then(
                (result) => {
                    resolve(result.ok);
                    // if (result.ok)
                    //     result.json().then((response) => { resolve(response); });
                },
                (error) => { reject(error); }
            );
        });
    },
};