/*
    @euisaquevenancio - 25/09/2026
    Automação para captura de relatórios no Vexpenses das mantenedoras ABEC, SOME, UBEE e UNBEC, desenvolvido com Node.js.
    
    Instalando todas bibliotecas de uma vez via terminal:
    npm install

    Instalando as bibliotecas manualmente via terminal:
    npm install dotenv
    npm install axios@1.4.0 cheerio@1.0.0-rc.12
    npm install puppeteer
    npm install exceljs

    Executando o código via terminal:
    node autoVexpenses.js
*/

// Bibliotecas utilizadas
require("dotenv").config(); // Manipulação de variáveis de ambiente .env
const fs = require("fs"); // Manipulação de arquivos - File System
const puppeteer = require("puppeteer"); // Manipulação do navegador - Chrome ou Firefox
const ExcelJS = require("exceljs"); // Manipulação de arquivos Excel
const path = require("path"); // Manipulação de caminhos de arquivos
const { exec } = require("child_process"); // Execução de comandos do sistema operacional - necessário para abrir o arquivo Excel no final do processo

const email = process.env.EMAIL;
const senha = process.env.SENHA;

const listaRelatoriosIgnorados = fs.readFileSync("dados/relatoriosIgnorados.txt", "utf-8")
                                .split("\n")
                                .map((relatorioIgnorado) => relatorioIgnorado?.trim())
                                .filter((relatorioIgnorado) => relatorioIgnorado.length > 0);
let relatoriosDasPaginas = [];

async function main() {
    // Declarando o navegador
    const navegador = await puppeteer.launch({
        executablePath: "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
        headless: false // Visualizar ou não a automação rodando
    });
    
    const horarioInicio = new Date().toLocaleTimeString("pt-BR");
    
    const abaNavegador = await navegador.newPage();
    abaNavegador.setDefaultTimeout(30000);

    await abaNavegador.goto(
        "https://auth.vexpenses.com/login",
        { waitUntil: "networkidle2" }
    );

    // Realiza login no Vexpenses
    if (!(await abaNavegador.$("#email") !== null)) {
        console.log("📢 O campo email não foi encontrado!");
        return false;
    }

    await abaNavegador.type("#email", email);
    await abaNavegador.keyboard.press("Enter");

    await abaNavegador.waitForNavigation({ waitUntil: "networkidle2" });
    await new Promise((r) => setTimeout(r, 7000));

    if (!(await abaNavegador.$("#i0116") !== null)) {
        console.log("📢 O campo email não foi encontrado!");
        return false;
    }

    await abaNavegador.type("#i0116", email);
    await abaNavegador.keyboard.press("Enter");
    
    await new Promise((r) => setTimeout(r, 7000));

    if (!(await abaNavegador.$("#i0118") !== null)) {
        console.log("📢 O campo senha não foi encontrado!");
        return false;
    }

    await abaNavegador.type("#i0118", senha);
    await abaNavegador.keyboard.press("Enter");

    await new Promise((r) => setTimeout(r, 8000));
    
    await abaNavegador.click("#idSIButton9", { clickCount: 1 });

    await abaNavegador.waitForNavigation({ waitUntil: "networkidle2" });
    await new Promise((r) => setTimeout(r, 8000));

    // Links de reembolso, adiantamento e prestação de contas
    const links = [
        { tipoDespesa: "REEMBOLSO", link: "https://amp.vexpenses.com/aprovacoes?approvals_filters=%22%7B%5C%22state%5C%22%3A%7B%5C%22costsCenters%5C%22%3Anull%2C%5C%22project%5C%22%3Anull%2C%5C%22page%5C%22%3A1%2C%5C%22perPage%5C%22%3A50%2C%5C%22members%5C%22%3Anull%2C%5C%22approvalFlows%5C%22%3Anull%2C%5C%22search%5C%22%3A%5C%22%5C%22%2C%5C%22statusTab%5C%22%3A%5C%22approved%5C%22%2C%5C%22orderBy%5C%22%3A%5C%22send_date%5C%22%2C%5C%22referenceDate%5C%22%3A%5C%22send_date%5C%22%2C%5C%22coin%5C%22%3Anull%2C%5C%22paymentsMethods%5C%22%3A%5B%7B%5C%22label%5C%22%3A%5C%22Reembolso%5C%22%2C%5C%22value%5C%22%3A639919%7D%2C%7B%5C%22label%5C%22%3A%5C%22Reembolso+Gastos+Eventuais%5C%22%2C%5C%22value%5C%22%3A659335%7D%5D%2C%5C%22order%5C%22%3A%5C%22desc%5C%22%7D%2C%5C%22version%5C%22%3A0%7D%22"},
        { tipoDespesa: "ADIANTAMENTO", link: "https://amp.vexpenses.com/aprovacoes?approvals_filters=%22%7B%5C%22state%5C%22%3A%7B%5C%22costsCenters%5C%22%3Anull%2C%5C%22project%5C%22%3Anull%2C%5C%22page%5C%22%3A1%2C%5C%22perPage%5C%22%3A50%2C%5C%22members%5C%22%3Anull%2C%5C%22approvalFlows%5C%22%3Anull%2C%5C%22search%5C%22%3A%5C%22%5C%22%2C%5C%22statusTab%5C%22%3A%5C%22approved%5C%22%2C%5C%22orderBy%5C%22%3A%5C%22send_date%5C%22%2C%5C%22referenceDate%5C%22%3A%5C%22send_date%5C%22%2C%5C%22coin%5C%22%3Anull%2C%5C%22paymentsMethods%5C%22%3A%5B%7B%5C%22label%5C%22%3A%5C%22Adiantamento%5C%22%2C%5C%22value%5C%22%3A640034%7D%5D%2C%5C%22order%5C%22%3A%5C%22desc%5C%22%7D%2C%5C%22version%5C%22%3A0%7D%22"},
        { tipoDespesa: "PRESTAÇÃO DE CONTAS", link: "https://amp.vexpenses.com/aprovacoes?approvals_filters=%22%7B%5C%22state%5C%22%3A%7B%5C%22costsCenters%5C%22%3Anull%2C%5C%22project%5C%22%3Anull%2C%5C%22page%5C%22%3A1%2C%5C%22perPage%5C%22%3A50%2C%5C%22members%5C%22%3Anull%2C%5C%22approvalFlows%5C%22%3Anull%2C%5C%22search%5C%22%3A%5C%22%5C%22%2C%5C%22statusTab%5C%22%3A%5C%22approved%5C%22%2C%5C%22orderBy%5C%22%3A%5C%22send_date%5C%22%2C%5C%22referenceDate%5C%22%3A%5C%22send_date%5C%22%2C%5C%22coin%5C%22%3Anull%2C%5C%22paymentsMethods%5C%22%3A%5B%7B%5C%22label%5C%22%3A%5C%22Presta%C3%A7%C3%A3o+de+Contas+do+Adiantamento%5C%22%2C%5C%22value%5C%22%3A652193%7D%2C%7B%5C%22label%5C%22%3A%5C%22Presta%C3%A7%C3%A3o+de+Contas+Adiantamento+-+Moeda+Estrangeira%5C%22%2C%5C%22value%5C%22%3A654665%7D%5D%2C%5C%22order%5C%22%3A%5C%22desc%5C%22%7D%2C%5C%22version%5C%22%3A0%7D%22"}
    ];

    for (let i = 0; i < links.length; i++) {
        // Se o login der certo, vai para a página de aprovações (relatórios que devem ser validados)
        await abaNavegador.goto(
            links[i].link,
            { waitUntil: "networkidle2" }
        );
        await new Promise((r) => setTimeout(r, 8000));

        if (i == 0) {
            await abaNavegador.click("#adopt-accept-all-button", { clickCount: 1 });
            await abaNavegador.click(".sc-elfQTZ.kYfrkC svg", { clickCount: 1 });
        }

        let primeiraVez = true;

        await new Promise((r) => setTimeout(r, 8000));

        do {
            const relatoriosDaPagina = await abaNavegador.$$(".approval-table-row");

            for (let j = 0; j < relatoriosDaPagina.length; j++) {
                let numeroRelatorio = await relatoriosDaPagina[j].$$eval(".sc-fYiONg.gesIpF.td-with-link-truncate a p", elementos => elementos[1].textContent?.trim());
                numeroRelatorio = numeroRelatorio.replace(/^#/, "").split(" - ")[0];
                
                if (listaRelatoriosIgnorados.includes(numeroRelatorio)) continue;
                
                const nomeSolicitante = await relatoriosDaPagina[j].$$eval(".sc-fYiONg.gesIpF.td-with-link-truncate a p", elementos => elementos[0].textContent?.trim());
                const valorRelatorio = await relatoriosDaPagina[j].$$eval(".sc-fYiONg.gesIpF.td-with-link-truncate a p", elementos => elementos[4].textContent?.trim());

                relatoriosDasPaginas.push({
                    numeroRelatorio: numeroRelatorio,
                    nomeSolicitante: nomeSolicitante,
                    valorRelatorio: valorRelatorio,
                    tipoDespesa: links[i].tipoDespesa
                });
            }

            await abaNavegador.waitForSelector(".sc-jKYigd.czRWRs svg", {
                timeout: 8000
            });

            const botoesSvg = await abaNavegador.$$(".sc-jKYigd.czRWRs svg");

            // Primeira vez: executa mesmo tendo apenas 1 botão
            if (primeiraVez) {
                if (botoesSvg.length == 0) {
                    break;
                }

                primeiraVez = false;
                await botoesSvg[0].click();
                await new Promise((r) => setTimeout(r, 8000));
                continue;
            }

            // Depois da primeira vez:
            if (botoesSvg.length === 2) {
                await botoesSvg[1].click();
                await new Promise((r) => setTimeout(r, 8000));
                continue;
            }

            break;
        } while (true);
    }

    console.log();
    console.log(`NOVOS RELATÓRIOS: ${relatoriosDasPaginas.length}`);
    console.log();

    for (let i = 0; i < relatoriosDasPaginas.length; i++) {
        // Acessando cada relatório
        await abaNavegador.goto(
            "https://amp.vexpenses.com/aprovacoes/" + relatoriosDasPaginas[i].numeroRelatorio,
            { waitUntil: "networkidle2" }
        );
        await new Promise((r) => setTimeout(r, 8000));

        const dataVencimento = await abaNavegador.$eval(".sc-jkdSes.iGToLf p, .sc-dpHMYC.deQJWY p, .sc-jonSUS.ipxdnJ p", el => el.textContent);
        if (relatoriosDasPaginas[i].tipoDespesa == "REEMBOLSO" || relatoriosDasPaginas[i].tipoDespesa == "ADIANTAMENTO") {
            relatoriosDasPaginas[i].dataVencimento = dataVencimento.trim().match(/\d{2}\/\d{2}/)?.[0] + "/2026";
        } else {
            relatoriosDasPaginas[i].dataVencimento = gerarDataVencimentoPrestacaoContas();
        }

        const botaoExibirHistorico = await abaNavegador.$(".sc-iBIXdp.fIUrel.history-button");
        await botaoExibirHistorico.click();

        await new Promise((r) => setTimeout(r, 8000));

        const mensagensHistorico = await abaNavegador.$$(".sc-hihqeJ.fTtTMe.timeline-item-container");

        const dataAprovacaoGestor = await mensagensHistorico[mensagensHistorico.length - 2].$eval(".sc-eZFbIh.kbOpOv span p", elemento => elemento.textContent?.trim().split(" • ")[0]);
        relatoriosDasPaginas[i].dataAprovacaoGestor = dataAprovacaoGestor;

        const centroCustos = await abaNavegador.$eval(".sc-UqsR.fPIsQV.cost-center-value-td-cell.custom-color", el => el.textContent.trim().split(" - ")[0]);
        relatoriosDasPaginas[i].centroCustos = centroCustos;
        
        const botaoExibirMaisDados = await abaNavegador.$("#table-pagination-container > div > table > tbody > tr:nth-child(1)");
        await botaoExibirMaisDados.evaluate(elemento => elemento.click());

        await abaNavegador.waitForSelector(".sc-bjbDXv.iJEdUH", {
            timeout: 8000
        });

        const elementoMantenedora = await abaNavegador.$(".sc-bjbDXv.iJEdUH");
        let mantenedora = await elementoMantenedora.evaluate(
            elemento => elemento.textContent?.trim() || ""
        );

        if (mantenedora[0] == 1) {
            mantenedora = "1 ABEC";
        } else if (mantenedora[0] == 2) {
            mantenedora = "2 SOME";
        } else if (mantenedora[0] == 3) {
            mantenedora = "3 UBEE";
        } else if (mantenedora[0] == 4) {
            mantenedora = "4 UNBEC";
        }
        relatoriosDasPaginas[i].mantenedora = mantenedora;

        console.log(`✅ #${(i+1)} | ${relatoriosDasPaginas[i].numeroRelatorio} | ${relatoriosDasPaginas[i].mantenedora} | ${relatoriosDasPaginas[i].nomeSolicitante} | ${relatoriosDasPaginas[i].valorRelatorio} | ${relatoriosDasPaginas[i].dataAprovacaoGestor} | ${relatoriosDasPaginas[i].dataVencimento} | ${relatoriosDasPaginas[i].tipoDespesa} | ${relatoriosDasPaginas[i].centroCustos}}`);
    }

    const horarioFim = new Date().toLocaleTimeString("pt-BR");
    console.log(`\n🤖 Fim da execução do script às ${horarioFim}.`);
    console.log(`🕓 Tempo de execução: ${calcularDiferencaHoras(horarioInicio, horarioFim)}.\n`);

    await salvarRelatorios();
    navegador.close();
}

function gerarDataVencimentoPrestacaoContas() {
    const data = new Date();
    let diasUteis = 0;

    while (diasUteis < 7) {
        data.setDate(data.getDate() + 1);

        const diaSemana = data.getDay();

        if (diaSemana !== 0 && diaSemana !== 6) {
            diasUteis++;
        }
    }

    return data.toLocaleDateString("pt-BR");
}

async function salvarRelatorios() {
    const pastaDestino = path.join(__dirname, "dados");

    // Garante que a pasta exista
    if (!fs.existsSync(pastaDestino)) {
        fs.mkdirSync(pastaDestino, { recursive: true });
    }

    // Captura o arquivo Excel
    const arquivoExcel = path.join(pastaDestino, "relatorios.xlsx");

    const workbook = new ExcelJS.Workbook();
    // Se o arquivo existe, lê
    if (fs.existsSync(arquivoExcel)) {
        await workbook.xlsx.readFile(arquivoExcel);
    }

    let planilha = workbook.getWorksheet("Relatorios");
    // Se a planilha não estiver estruturada, adiciona o cabeçalho
    if (!planilha) {
        planilha = workbook.addWorksheet("Relatorios");

        planilha.addRow([
            "N° RELATÓRIO (VEXPESES)",
            "MANTENEDORA",
            "NOME",
            "VALOR",
            "DATA APROVAÇÃO GESTOR",
            "DATA VENCIMENTO",
            "TIPO DE DESPESA",
            "STATUS",
            "CR (SOMENTE PARA UBEE E UNBEC)"
        ]);
    }

    // Aplica a largura nas colunas
    planilha.columns = [
        { key: "N° RELATÓRIO (VEXPESES)", width: 8 },
        { key: "MANTENEDORA", width: 8 },
        { key: "NOME", width: 8 },
        { key: "VALOR", width: 8 },
        { key: "DATA APROVAÇÃO GESTOR", width: 8 },
        { key: "DATA VENCIMENTO", width: 8 },
        { key: "TIPO DE DESPESA", width: 8 },
        { key: "STATUS", width: 8 },
        { key: "CR (SOMENTE PARA UBEE E UNBEC)", width: 8 }
    ];

    // Adicionando os tickets na planilha
    if (relatoriosDasPaginas.length > 0) {
        for (let i = 0; i < relatoriosDasPaginas.length; i++) {
            const dataAprovacaoGestor = converterParaData(relatoriosDasPaginas[i].dataAprovacaoGestor);
            const dataVencimento = converterParaData(relatoriosDasPaginas[i].dataVencimento);

            planilha.addRow([
                relatoriosDasPaginas[i].numeroRelatorio,
                relatoriosDasPaginas[i].mantenedora,
                relatoriosDasPaginas[i].nomeSolicitante,
                relatoriosDasPaginas[i].valorRelatorio.substring(3),
                dataAprovacaoGestor,
                dataVencimento,
                relatoriosDasPaginas[i].tipoDespesa,
                "LANÇAR",
                relatoriosDasPaginas[i].centroCustos
            ]);
        }
    }

    if (planilha.rowCount > 1) {
        // Remove os relatórios da tabela antiga
        if (planilha.model.tables) {
            planilha.model.tables = [];
        }

        const linhasValidas = planilha.getSheetValues().slice(2) // Remove o cabeçalho
                                                       .filter(linha => Array.isArray(linha)) // Remove undefined
                                                       .map(linha => linha.slice(1)); // Remove índice fantasma
        
        planilha.addTable({
            name: "TabelaRelatorios",
            ref: "A1",
            headerRow: true,
            style: {
                theme: "TableStyleLight1"
            },
            columns: [
                { name: "N° RELATÓRIO (VEXPESES)" },
                { name: "MANTENEDORA" },
                { name: "NOME" },
                { name: "VALOR" },
                { name: "DATA APROVAÇÃO GESTOR" },
                { name: "DATA VENCIMENTO" },
                { name: "TIPO DE DESPESA" },
                { name: "STATUS" },
                { name: "CR (SOMENTE PARA UBEE E UNBEC)" }
            ],
            rows: linhasValidas
        });
    }

    if (relatoriosDasPaginas.length > 0) {
        // Salva o arquivo excel
        await workbook.xlsx.writeFile(arquivoExcel);

        // Adicionando os novos relatórios que devem ser ignorados (já estão na planilha)
        const arquivoRelatoriosIgnorados = path.join(pastaDestino, "relatoriosIgnorados.txt");

        let conteudoExistenteRelatoriosIgnorados = "";
        // Se o arquivo existir, lê o conteúdo
        if (fs.existsSync(arquivoRelatoriosIgnorados)) {
            conteudoExistenteRelatoriosIgnorados = fs.readFileSync(arquivoRelatoriosIgnorados, "utf-8");
        }

        let novoConteudoRelatoriosIgnorados = "";
        for (const relatorioAtual of relatoriosDasPaginas) {
            novoConteudoRelatoriosIgnorados += `\n${relatorioAtual.numeroRelatorio}`;
        }

        // Só escreve no TXT se tiver conteúdo novo
        if (novoConteudoRelatoriosIgnorados != "") {
            fs.appendFileSync(arquivoRelatoriosIgnorados, novoConteudoRelatoriosIgnorados, "utf-8");
        }

        exec(`start "" "${arquivoExcel}"`);
    }
}

// Função para calcular a diferença entre dois horários no formato HH:mm:ss
function calcularDiferencaHoras(horarioInicio, horarioFim) {
    try {
        // Quebra as strings em partes
        const [h1, m1, s1] = horarioInicio.split(":").map(Number);
        const [h2, m2, s2] = horarioFim.split(":").map(Number);

        // Cria objetos Date no mesmo dia
        const dataBase = new Date();
        const date1 = new Date(dataBase.getFullYear(), dataBase.getMonth(), dataBase.getDate(), h1, m1, s1 || 0);
        const date2 = new Date(dataBase.getFullYear(), dataBase.getMonth(), dataBase.getDate(), h2, m2, s2 || 0);

        // Calcula a diferença em milissegundos
        let diffMs = date2 - date1;

        // Se negativo, inverte
        const negativo = diffMs < 0;
        diffMs = Math.abs(diffMs);

        // Converte para horas, minutos e segundos
        const horas = Math.floor(diffMs / (1000 * 60 * 60));
        const minutos = Math.floor((diffMs % (1000 * 60 * 60)) / (1000 * 60));
        const segundos = Math.floor((diffMs % (1000 * 60)) / 1000);

        return `${negativo ? '-' : ''}${String(horas).padStart(2, '0')}:${String(minutos).padStart(2, '0')}:${String(segundos).padStart(2, '0')}`;
    } catch (err) {
        console.error("Erro ao calcular diferença de horas: ", err);
        return null;
    }
}

function converterParaData(data) {
    if (!data) return null;

    // Se já for um objeto Date
    if (data instanceof Date) {
        return data;
    }

    // Se vier como string no formato dd/mm/aaaa
    if (typeof data === "string") {
        const partes = data.trim().split("/");

        if (partes.length === 3) {
            const dia = Number(partes[0]);
            const mes = Number(partes[1]);
            const ano = Number(partes[2]);

            return new Date(ano, mes - 1, dia);
        }
    }

    return null;
}

// Executando o código
main().catch((err) => {
    console.error("Erro na execução do script: ", err);
});