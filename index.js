const {Builder, By, WebDriver, until} = require('selenium-webdriver');
const planilhaChallenge = 'Files/challenge.xlsx'
const excelJson = require('convert-excel-to-json');

let rpaTeste = (async function rpachallengeFunction(){
        try {
                // conversao da planilha em json 
                const challengePlanilhaJson = excelJson({
                        sourceFile: planilhaChallenge,
                        header:{
                                rows: 1
                        },
                        columnToKey: {
                                '*': "{{columnHeader}}"
                        },
                });
                const challengeInput = challengePlanilhaJson.Sheet1;

                // Criacao do driver para inicializacao do chrome 
                let driver =  await new Builder().forBrowser('chrome').build();
                await driver.get('http://rpachallenge.com/'); // aguarda carregamento da pagina 

                //Btn start challenge
                await new Promise(r => setTimeout(r, 500));
                const btnIniciar = await driver.findElement(By.xpath('//button[text()="Start"]')); 
                btnIniciar.click();

                // Iteracao json extraido da planilhja 
                for (let index = 0; index < challengeInput.length; index++) {
                         // Parametrizacao de informações
                        const row = challengeInput[index];
                        const campoFirstName = row['First Name'];
                        const campoLastName = row['Last Name'];
                        const campoCompanyName = row['Company Name'];
                        const campoRoleInCompany = row['Role in Company'];
                        const campoAddress = row['Address'];
                        const campoEmail = row['Email'];
                        const campoPhoneNumber = row['Phone Number'];

                        // Buscando o formulário
                        const divInput = await driver.findElement(By.className('inputFields'));
                        // Buscando btn submit
                        const buttonSubmit = await driver.findElement(By.className("btn uiColorButton"));

                        // Buscando inputs a serem preenchidos
                        const inputFirstName = divInput.findElement(By.xpath("//rpa1-field[@ng-reflect-dictionary-value='First Name']//input"));
                        const inputLastName = divInput.findElement(By.xpath("//rpa1-field[@ng-reflect-dictionary-value='Last Name']//input"));
                        const inputCompanyName = divInput.findElement(By.xpath("//rpa1-field[@ng-reflect-dictionary-value='Company Name']//input"));
                        const inputRoleInCompany = divInput.findElement(By.xpath("//rpa1-field[@ng-reflect-dictionary-value='Role in Company']//input"));
                        const inputAddress = divInput.findElement(By.xpath("//rpa1-field[@ng-reflect-dictionary-value='Address']//input"));
                        const inputEmail = divInput.findElement(By.xpath("//rpa1-field[@ng-reflect-dictionary-value='Email']//input"));
                        const inputPhoneNumber = divInput.findElement(By.xpath("//rpa1-field[@ng-reflect-dictionary-value='Phone Number']//input"));

                        // Limpando campos
                        inputFirstName.clear();
                        inputLastName.clear();
                        inputCompanyName.clear();
                        inputRoleInCompany.clear();
                        inputAddress.clear();
                        inputEmail.clear();
                        inputPhoneNumber.clear();

                        // Preenchendo os inputs
                        inputFirstName.sendKeys(campoFirstName);
                        inputLastName.sendKeys(campoLastName);
                        inputCompanyName.sendKeys(campoCompanyName);
                        inputRoleInCompany.sendKeys(campoRoleInCompany);
                        inputAddress.sendKeys(campoAddress);
                        inputEmail.sendKeys(campoEmail);
                        inputPhoneNumber.sendKeys(campoPhoneNumber);

                        // Delay para validacao do form
                        await new Promise(r => setTimeout(r, 300));
                        
                        // Envio do form
                        await buttonSubmit.click();
                }
        } catch (e) {
                await new Promise(r => setTimeout(r, 300));
                console.log(e);
        } finally {
                console.log('Teste finalizado');
                await driver.quit();
        }
});

rpaTeste();
