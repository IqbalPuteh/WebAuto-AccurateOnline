using Microsoft.Playwright;
using System.Diagnostics;
using System.Configuration;
using System.IO.Compression;
using System.Threading;
using System;
using System.IO;

internal class program
{
    // Add a logger to a file
    private static TextWriterTraceListener logListener;

    // Define the app downlad folder name in 'Download' folder
    static String strAppDownloadFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\" + ConfigurationManager.AppSettings["AppDownloadFolder"];
    static String strDataSharingFolder = strAppDownloadFolder + @"\" + ConfigurationManager.AppSettings["DataFolder"];
    static String strUploadFolder = strAppDownloadFolder + @"\" + ConfigurationManager.AppSettings["UploadFolder"];
    static String strLogsFolder = strAppDownloadFolder + @"\" + ConfigurationManager.AppSettings["LogFolderName"];

    // Define the logger to console
    static String enableConsoleLogging = ConfigurationManager.AppSettings["ConsoleScreenLogMessage"];
    // Defince the trace element logger to console
    static String enableElementCheckLogging = ConfigurationManager.AppSettings["TraceElement"];

    // Define the sandbox setting
    static String strSandBoxBool = ConfigurationManager.AppSettings["ConnectToSandbox"];

    // Define the DT Id & Name
    static String strDTid = ConfigurationManager.AppSettings["DT-Id"];
    static String strDTname = ConfigurationManager.AppSettings["DT-Name"];

    // Define debug browser -- headless=> "Y" (or True) flag is for visible
    static String strBrowserVisible = ConfigurationManager.AppSettings["BrowserIsVisible"];

    // Define Accounting data flag
    static String strAcctDataIncl = ConfigurationManager.AppSettings["IncludeAccountingData"];

    // Flag for all page simulation process is success
    static Boolean allpageprocessSuccess = false;

    static string logFilePathAndName = "";

    // Entry point of the program
    static async Task Main()
    {
        // Initialize the logger
        InitializeLogger(enableConsoleLogging);

        try
        {
            // Log program start
            Log("###  Program started  ###");
            Log("");

            // Log time stamp
            Log("Time Stamp: " + DateTime.Now.ToString("dd MMM yyyy HH:mm:ss"));

            // Step 1: Create/Check folder and Delete supporting files
            DeleteAllSupportingFiles();
            Log("Supporting files deleted.");

            // Step 2: Open Chrome and login to a web application
            int retryCount = 3;
            bool success = false;

            for (int i = 0; i < retryCount; i++)
            {
                try
                {
                    success = await OpenBrowserAndLogin();
                    if (success)
                    {
                        Log("Web automation successful");
                        break;
                    }
                    else { Thread.Sleep(30000); }
                }
                catch (Exception ex)
                {
                    Log($"Attempt {i + 1} failed with error: {ex.Message}");
                    if (i == retryCount - 1)
                    {
                        Log("All attempts failed");
                    }
                }
            }

            // Step 3: Zip and send files
            if (success) { ZipAndSendFile(); }
        
            // Log program end
            Log("");
            Log("###  Program finished ###");
        }
        catch (Exception ex)
        {
            // Log any exceptions that occur
            Log($"An error occurred: {ex.Message}");
        }
        finally
        {
            // Close the logger
            CloseLogger();
        }
    }

    static string GetPrevMonth()
    {
        return DateTime.Now.AddMonths(-1).ToString("MM");
    }

    static string GetPrevYear()
    {
        return DateTime.Now.AddMonths(-1).ToString("yyyy");
    }

    static string GetDSPeriod()
    {
        return GetPrevYear() + GetPrevMonth();
    }

    static string GetFirstDate()
    {
        return "01";
    }

    static string GetLastDayOfPrevMonth()
    {
        var lastDay = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddDays(-1);
        return lastDay.ToString("dd");
    }

    private static void InitializeLogger()
    {
        // Customize the log file path and name here
        string logFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads", "Debug-" +strDTid +"-"+ strDTname + ".log");

        // Create a new text writer for logging
        logListener = new TextWriterTraceListener(logFilePath);

        // Add the listener to the trace sources
        Trace.Listeners.Add(logListener);

        // Set the level of detail you want to log
        Trace.AutoFlush = true;
    }

    private static void InitializeLogger(string enableConsoleLogging = "N")
    {
        //Convert string var to boolean
        var boolConsoleLogging = false;
        if (enableConsoleLogging == "Y") { boolConsoleLogging = true; }

        // Customize the log file path and name here
        logFilePathAndName = Path.Combine(strAppDownloadFolder, "Debug-" + strDTid + "-" + strDTname + "-" + GetDSPeriod() + ".log");

        // Create a new text writer for logging to a file
        logListener = new TextWriterTraceListener(logFilePathAndName);

        // Add the listener to the trace sources
        Trace.Listeners.Add(logListener);

        // Optionally, add a console listener if writeToConsole is true
        if (boolConsoleLogging)
        {
            Trace.Listeners.Add(new ConsoleTraceListener());
        }

        // Set the level of detail you want to log
        Trace.AutoFlush = true;
    }


    // Close the logger when done
    private static void CloseLogger()
    {
        if (logListener != null)
        {
            logListener.Close();
            logListener.Dispose();
        }
    }

    // Step 2: Open Chrome and login to the web application

    private static async Task<Boolean> OpenBrowserAndLogin()
    {
        try
        {
            Log("");

            using var playwright = await Playwright.CreateAsync();
            IBrowser browser;

            if (strBrowserVisible == "Y")
            {
                browser = await playwright.Chromium.LaunchAsync(new()
                {
                    Headless = false,
                    SlowMo = 100
                });
            }
            else
            {
                browser = await playwright.Chromium.LaunchAsync(new()
                {
                    Headless = true,
                    SlowMo = 100
                });
            }

            var context = await browser.NewContextAsync();
            var page = await context.NewPageAsync();
            page.SetDefaultTimeout(30000);

            var dashboardUrl = "https://account.accurate.id/";

            await page.GotoAsync(dashboardUrl);
            Log($"Navigating to the dashboard URL: {dashboardUrl}");
            await page.WaitForLoadStateAsync(LoadState.NetworkIdle);

            Log("Filling in ID and password fields...");

            await page.GetByLabel("Email atau No Handphone").FillAsync(ConfigurationManager.AppSettings["LoginId"].ToString()) ;
            await page.GetByRole(AriaRole.Textbox, new() { Name = "Password" }).FillAsync(ConfigurationManager.AppSettings["Password"].ToString());

            await ScreenShotLog(page, 0);

            Log("Try to click on 'Masuk' button...");
            await page.ClickAsync("#btn-login");
            await page.WaitForLoadStateAsync(LoadState.NetworkIdle);
            await ScreenShotLog( page, 0);

            Log("Try to click on the PT Cipta Alam Bahagia container...");
            var attempts = 0;
            {
                while (attempts < 3)
                {
                    if (page.Url == "https://account.accurate.id/manage" && await IsElementExist("#content > div > div.col-lg-8.col-12.mt-1.manage-database > div > div > div > div > div.cursor-pointer > h3", page))
                    {
                        break;
                    }
                    else
                    {
                        attempts++;
                        if (attempts == 3)
                        {
                            await LogoutFunction(page);
                            return allpageprocessSuccess;
                        }
                    }
                    await Task.Delay(10000);
                }
                var newPage = await page.RunAndWaitForPopupAsync(async () =>
                {
                    await page.Locator("#content img").Nth(3).ClickAsync();
                });
                await newPage.WaitForLoadStateAsync();
                await ScreenShotLog( page, 15000);

                page = newPage;
                page.SetDefaultTimeout(35000); // ==>> do no change this
            }

            Log("Try to click on right-sided report icon menu / Daftar Laporan tab...");
            {
                await IsElementExist("i.main-menu-icon.icn-menu-mymenu", page);
                var element = await page.QuerySelectorAsync("i.main-menu-icon.icn-menu-mymenu");
                await element.ClickAsync();
                await page.WaitForTimeoutAsync(500);

                await page.GetByRole(AriaRole.Link, new() { Name = "Daftar Laporan" }).ClickAsync();
                await page.WaitForLoadStateAsync(LoadState.NetworkIdle);
                await ScreenShotLog(page, 10000);
            }

            Log("Confirming page is at correct URL address...");
            attempts = 0;
            while (attempts < 3)
            {
                if (page.Url.StartsWith("https://public.accurate.id/accurate/?_dsi") && await IsElementExist("div:nth-of-type(2) > div > div:nth-of-type(3) a", page))
                {
                    break;
                }
                else
                {
                    attempts++;
                    if (attempts == 3)
                    {
                        await LogoutFunction(page);
                        return allpageprocessSuccess;
                    }
                }
                await Task.Delay(10000);
            }


            Log("Try to click on 'Daftar Faktur Penjualan'...");
            if (await IsElementExist("div.index-report-tab-container > div:nth-of-type(2) > div > div:nth-of-type(3) a", page))
            {
                await page.Locator("div:nth-of-type(2) > div > div:nth-of-type(3) a").ClickAsync();

                //await page.WaitForLoadStateAsync(LoadState.NetworkIdle);
                await ScreenShotLog(page, 5000);

                Log("Try to send data parameters...");
                await IsElementExist("fieldset:nth-of-type(2) > div > div:nth-of-type(1) input", page);
                var element = await page.QuerySelectorAsync("fieldset:nth-of-type(2) > div > div:nth-of-type(1) input");
                await element.FillAsync(GetFirstDate() + GetPrevMonth() + GetPrevYear());
                await ScreenShotLog( page);

                await IsElementExist("fieldset:nth-of-type(2) > div > div:nth-of-type(2) input", page);
                element = await page.QuerySelectorAsync("fieldset:nth-of-type(2) > div > div:nth-of-type(2) input");
                await element.FillAsync(GetLastDayOfPrevMonth() + GetPrevMonth() + GetPrevYear());
                await ScreenShotLog(page);

                await IsElementExist("div:nth-of-type(9) div.text-right > button", page);
                await page.Locator("div:nth-of-type(9) div.text-right > button").ClickAsync();
                await ScreenShotLog(page, 10000);

               Log("Try to click on 'Download' button and waiting for download...");
                {
                    await IsElementExist("div.tab-control button.dropdown-toggle", page);
                    await page.Locator("div.tab-control button.dropdown-toggle").ClickAsync();
                    await ScreenShotLog(page, 5000);
                    await IsElementExist("#module-accurate__report__report li:nth-of-type(2) > a", page);
                    var waitForDownloadTask = page.WaitForDownloadAsync();
                    await page.Locator("#module-accurate__report__report li:nth-of-type(2) > a").ClickAsync();

                    var download = await waitForDownloadTask;
                    await download.SaveAsAsync($"{strAppDownloadFolder}/Sales_Data.xlsx");
                    Log($"Downloaded file path: {await download.PathAsync()}");
                }
            }

                // Back to daftar laporan
                Log("Heading back on 'Daftar Laporan' and try to click on Rincian Penerimaan Penjualan ...");
                if (await IsElementExist("div.module-tab > button > i", page))
                {
                    //Closing previouse report tab
                    var element = await page.QuerySelectorAsync("div.module-tab > div > button");
                    await element.ClickAsync();

                    element = await page.QuerySelectorAsync("div.module-tab > button > i");
                    await element.ClickAsync();
                } else
                {
                    await LogoutFunction(page);
                    return allpageprocessSuccess;
                }
            if (await IsElementExist("div.index-report-tab-container > div:nth-of-type(2) > div > div:nth-of-type(2) a", page))
            {

                await page.Locator("div.index-report-tab-container > div:nth-of-type(2) > div > div:nth-of-type(2) a").ClickAsync();
                await ScreenShotLog(page, 5000);

                Log("Sending data parameters...");
                await IsElementExist("fieldset:nth-of-type(2) > div > div:nth-of-type(1) input", page);
                var element = await page.QuerySelectorAsync("fieldset:nth-of-type(2) > div > div:nth-of-type(1) input");
                await element.FillAsync(GetFirstDate() + GetPrevMonth() + GetPrevYear());
                await ScreenShotLog(page);

                await IsElementExist("fieldset:nth-of-type(2) > div > div:nth-of-type(2) input", page);
                element = await page.QuerySelectorAsync("fieldset:nth-of-type(2) > div > div:nth-of-type(2) input");
                await element.FillAsync(GetLastDayOfPrevMonth() + GetPrevMonth() + GetPrevYear());
                await ScreenShotLog(page);

                await IsElementExist("div:nth-of-type(9) div.text-right > button", page);
                await page.Locator("div:nth-of-type(9) div.text-right > button").ClickAsync();
                await ScreenShotLog(page, 10000);
                Log("Try to click the 'Download' button and waiting for download...");
                {
                    await IsElementExist("div.tab-control button.dropdown-toggle", page);
                    await page.Locator("div.tab-control button.dropdown-toggle").ClickAsync();
                    await page.WaitForTimeoutAsync(5000);
                    await IsElementExist("#module-accurate__report__report li:nth-of-type(2) > a", page);
                    var waitForDownloadTask = page.WaitForDownloadAsync();
                    await page.Locator("#module-accurate__report__report li:nth-of-type(2) > a").ClickAsync();

                    var download = await waitForDownloadTask;
                    await download.SaveAsAsync($"{strAppDownloadFolder}/Repayment_Data.xlsx");
                    Log($"Downloaded file path: {await download.PathAsync()}");
                }
            }

            if (strAcctDataIncl == "Y")
            {
                // Begin downloading Accounting reports
                // Back to daftar laporan
                Log("Heading back on 'Daftar Laporan' and try to click on Laporan Arus Kas per Akun ...");
                if (await IsElementExist("div.module-tab > button > i", page))
                {
                    //Closing previouse report tab
                    var element = await page.QuerySelectorAsync("div.module-tab > div > button");
                    await element.ClickAsync();

                    element = await page.QuerySelectorAsync("div.module-tab > button > i");
                    await element.ClickAsync();
                }
                else
                {
                    await LogoutFunction(page);
                    return allpageprocessSuccess;
                }
                if (await IsElementExist("div.index-report-tab-container > div:nth-of-type(2) > div > div:nth-of-type(1) a", page))
                {
                    await page.Locator("div.index-report-tab-container > div:nth-of-type(2) > div > div:nth-of-type(1) a").ClickAsync();
                    await ScreenShotLog(page, 5000);

                    Log("Sending data parameters...");
                    await IsElementExist("fieldset:nth-of-type(2) > div > div:nth-of-type(1) input", page);
                    var element = await page.QuerySelectorAsync("fieldset:nth-of-type(2) > div > div:nth-of-type(1) input");
                    await element.FillAsync(GetFirstDate() + GetPrevMonth() + GetPrevYear());
                    await ScreenShotLog(page);

                    await IsElementExist("fieldset:nth-of-type(2) > div > div:nth-of-type(2) input", page);
                    element = await page.QuerySelectorAsync("fieldset:nth-of-type(2) > div > div:nth-of-type(2) input");
                    await element.FillAsync(GetLastDayOfPrevMonth() + GetPrevMonth() + GetPrevYear());
                    await ScreenShotLog(page);

                    await IsElementExist("fieldset:nth-of-type(9) button", page);
                    await page.Locator("fieldset:nth-of-type(9) button").ClickAsync();
                    await ScreenShotLog(page, 5000);

                    await page.GetByText("Kas Kecil").ClickAsync();

                    await IsElementExist("fieldset:nth-of-type(9) div > button", page);
                    await page.Locator("fieldset:nth-of-type(9) div > button").ClickAsync();
                    await page.WaitForTimeoutAsync(5000);

                    await page.GetByText("Bank", new() { Exact = true }).ClickAsync();

                    await IsElementExist("fieldset:nth-of-type(9) div > button", page);
                    await page.Locator("fieldset:nth-of-type(9) div > button").ClickAsync();
                    await page.WaitForTimeoutAsync(5000);

                    await page.GetByText("Deposito Bank").ClickAsync();
                    await ScreenShotLog(page);

                    await IsElementExist("div:nth-of-type(9) div.text-right > button", page);
                    await page.Locator("div:nth-of-type(9) div.text-right > button").ClickAsync();
                    await ScreenShotLog(page, 10000);
                    Log("Try to click the 'Download' button and waiting for download...");
                    {
                        await IsElementExist("div.tab-control button.dropdown-toggle", page);
                        await page.Locator("div.tab-control button.dropdown-toggle").ClickAsync();
                        await page.WaitForTimeoutAsync(5000);
                        await IsElementExist("#module-accurate__report__report li:nth-of-type(2) > a", page);
                        var waitForDownloadTask = page.WaitForDownloadAsync();
                        await page.Locator("#module-accurate__report__report li:nth-of-type(2) > a").ClickAsync();

                        var download = await waitForDownloadTask;
                        await download.SaveAsAsync($"{strAppDownloadFolder}/Arus_Kas.xlsx");
                        Log($"Downloaded file path: {await download.PathAsync()}");
                    }
                }

                // Back to daftar laporan
                Log("Heading back on 'Daftar Laporan' and try to click on Ringkasan Mutasi Gudang1 ...");
                if (await IsElementExist("div.module-tab > button > i", page))
                {
                    //Closing previouse report tab
                    var element = await page.QuerySelectorAsync("div.module-tab > div > button");

                    await element.ClickAsync();

                    element = await page.QuerySelectorAsync("div.module-tab > button > i");
                    await element.ClickAsync();
                } else
                {
                    await LogoutFunction(page);
                    return allpageprocessSuccess;
                }
                if (await IsElementExist("div.index-report-tab-container > div:nth-of-type(2) > div > div:nth-of-type(4) a", page))
                {
                    await page.Locator("div.index-report-tab-container > div:nth-of-type(2) > div > div:nth-of-type(4) a").ClickAsync();

                    await ScreenShotLog(page, 5000);

                    Log("Sending data parameters...");
                    await IsElementExist("fieldset:nth-of-type(2) > div > div:nth-of-type(1) input", page);
                    var element = await page.QuerySelectorAsync("fieldset:nth-of-type(2) > div > div:nth-of-type(1) input");
                    await element.FillAsync(GetFirstDate() + GetPrevMonth() + GetPrevYear());
                    await ScreenShotLog(page);

                    await IsElementExist("fieldset:nth-of-type(2) > div > div:nth-of-type(2) input", page);
                    element = await page.QuerySelectorAsync("fieldset:nth-of-type(2) > div > div:nth-of-type(2) input");
                    await element.FillAsync(GetLastDayOfPrevMonth() + GetPrevMonth() + GetPrevYear());
                    await ScreenShotLog(page);

                    await IsElementExist("fieldset:nth-of-type(9) button", page);
                    await page.Locator("fieldset:nth-of-type(9) button").ClickAsync();
                    await ScreenShotLog(page, 5000);
                    await page.GetByText("Pusat").ClickAsync();
                    await ScreenShotLog(page);

                    await IsElementExist("div:nth-of-type(9) div.text-right > button", page);
                    await page.Locator("div:nth-of-type(9) div.text-right > button").ClickAsync();
                    await ScreenShotLog(page, 10000);
                    Log("Try to click the 'Download' button and waiting for download...");
                    {
                        await IsElementExist("div.tab-control button.dropdown-toggle", page);
                        await page.Locator("div.tab-control button.dropdown-toggle").ClickAsync();
                        await page.WaitForTimeoutAsync(5000);
                        await IsElementExist("#module-accurate__report__report li:nth-of-type(2) > a", page);
                        var waitForDownloadTask = page.WaitForDownloadAsync();
                        await page.Locator("#module-accurate__report__report li:nth-of-type(2) > a").ClickAsync();

                        var download = await waitForDownloadTask;
                        await download.SaveAsAsync($"{strAppDownloadFolder}/Inventory.xlsx");
                        Log($"Downloaded file path: {await download.PathAsync()}");
                    }
                }

                // Back to daftar laporan
                Log("");
                Log("Heading back on 'Daftar Laporan' and try to get to Keuangan list report tab....");
                if (await IsElementExist("div.module-tab > button > i", page))
                {
                    //Closing previouse report tab
                    var element = await page.QuerySelectorAsync("div.module-tab > div > button");
                    await element.ClickAsync();

                    element = await page.QuerySelectorAsync("div.module-tab > button > i");
                    await element.ClickAsync();

                } else
                {
                    await LogoutFunction(page);
                    return allpageprocessSuccess;
                }
                //Changing tab to Keuangan list report tab
                await page.Locator("span").Filter(new() { HasText = "Keuangan" }).ClickAsync();
                //Changing report tab done

                Log("Try to click on Laba/Rugi Standar...");
                if (await IsElementExist("div.index-report-tab-container > div:nth-of-type(2) > div > div:nth-of-type(1) a", page))
                {
                    await page.Locator("div.index-report-tab-container > div:nth-of-type(2) > div > div:nth-of-type(1) a").ClickAsync();
                    await ScreenShotLog(page, 5000);

                    Log("Sending data parameters...");
                    await IsElementExist("input[name=\"startDate\"]", page);
                    var element = await page.QuerySelectorAsync("input[name=\"startDate\"]");
                    await element.FillAsync(GetFirstDate() + GetPrevMonth() + GetPrevYear());
                    await ScreenShotLog(page);

                    await IsElementExist("input[name=\"endDate\"]", page);
                    element = await page.QuerySelectorAsync("input[name=\"endDate\"]");
                    await element.FillAsync(GetLastDayOfPrevMonth() + GetPrevMonth() + GetPrevYear());
                    await ScreenShotLog(page);

                    await page.Locator("label").Filter(new() { HasText = "Tampilkan Akun Induk" }).Locator("span").First.ClickAsync();
                    await page.Locator("label").Filter(new() { HasText = "Tampilkan data dengan Saldo Nol" }).Locator("span").First.ClickAsync();
                    await page.Locator("label").Filter(new() { HasText = "Tampilkan Saldo Akun Induk" }).Locator("span").First.ClickAsync();
                    await ScreenShotLog(page);

                    await IsElementExist("div:nth-of-type(9) div.text-right > button", page);
                    await page.Locator("div:nth-of-type(9) div.text-right > button").ClickAsync();
                    await ScreenShotLog(page, 10000);
                    Log("Try to click the 'Download' button and waiting for download...");
                    {
                        await IsElementExist("div.tab-control button.dropdown-toggle", page);
                        await page.Locator("div.tab-control button.dropdown-toggle").ClickAsync();
                        await page.WaitForTimeoutAsync(5000);
                        await IsElementExist("#module-accurate__report__report li:nth-of-type(2) > a", page);
                        var waitForDownloadTask = page.WaitForDownloadAsync();
                        await page.Locator("#module-accurate__report__report li:nth-of-type(2) > a").ClickAsync();

                        var download = await waitForDownloadTask;
                        await download.SaveAsAsync($"{strAppDownloadFolder}/Laporan_Laba_Rugi.xlsx");
                        Log($"Downloaded file path: {await download.PathAsync()}");
                    }
                }

                // Back to daftar laporan
                Log("Heading back on 'Daftar Laporan' and try to get to Keuangan list report tab....");
                if (await IsElementExist("div.module-tab > button > i", page))
                {
                    //Closing previouse report tab
                    var element = await page.QuerySelectorAsync("div.module-tab > div > button");
                    await element.ClickAsync();

                    element = await page.QuerySelectorAsync("div.module-tab > button > i");
                    await element.ClickAsync();
                }
                else
                {
                    await LogoutFunction(page);
                    return allpageprocessSuccess;
                }
                Log("Try to click on Neraca Standar...");
                if (await IsElementExist("div.index-report-tab-container > div:nth-of-type(2) > div > div:nth-of-type(2) a", page))
                {
                    await page.GetByRole(AriaRole.Link, new() { Name = "Neraca (Standar)" }).ClickAsync();

                    await IsElementExist("input[name=\"asOfDate\"]", page);
                    var element = await page.QuerySelectorAsync("input[name=\"asOfDate\"]");
                    await element.FillAsync(GetLastDayOfPrevMonth() + GetPrevMonth() + GetPrevYear());
                    await ScreenShotLog(page);

                    await page.Locator("label").Filter(new() { HasText = "Tampilkan Akun Induk" }).Locator("span").First.ClickAsync();
                    await page.Locator("label").Filter(new() { HasText = "Tampilkan data dengan Saldo Nol" }).Locator("span").First.ClickAsync();
                    await page.Locator("label").Filter(new() { HasText = "Tampilkan Saldo Akun Induk" }).Locator("span").First.ClickAsync();

                    await IsElementExist("div:nth-of-type(9) div.text-right > button", page);
                    await page.Locator("div:nth-of-type(9) div.text-right > button").ClickAsync();
                    await ScreenShotLog(page, 10000);
                    Log("Try to click the 'Download' button and waiting for download...");
                    {
                        await IsElementExist("div.tab-control button.dropdown-toggle", page);
                        await page.Locator("div.tab-control button.dropdown-toggle").ClickAsync();
                        await page.WaitForTimeoutAsync(5000);
                        await IsElementExist("#module-accurate__report__report li:nth-of-type(2) > a", page);
                        var waitForDownloadTask = page.WaitForDownloadAsync();
                        await page.Locator("#module-accurate__report__report li:nth-of-type(2) > a").ClickAsync();

                        var download = await waitForDownloadTask;
                        await download.SaveAsAsync($"{strAppDownloadFolder}/Neraca_Saldo.xlsx");
                        Log($"Downloaded file path: {await download.PathAsync()}");
                    }
                }
            }

            // Back to daftar laporan
            Log("");
            Log("Heading back on 'Daftar Laporan' and try to get to Piutang list report tab....");
            if (await IsElementExist("div.module-tab > button > i", page))
            {
                //Closing previous report tab
                var element = await page.QuerySelectorAsync("div.module-tab > div > button");
                await element.ClickAsync();

                element = await page.QuerySelectorAsync("div.module-tab > button > i");
                await element.ClickAsync();

            }
            else
            {
                await LogoutFunction(page);
                return allpageprocessSuccess;
            }
            //Changing tab to Piutang list report tab
            await page.Locator("span").Filter(new() { HasText = "Piutang" }).ClickAsync();
            //Changing report tab done

            Log("Try to click on Daftar Pelanggan...");
            if (await IsElementExist("div.index-report-tab-container > div:nth-of-type(2) > div > div:nth-of-type(2) a", page))
            {
                await page.Locator("div.index-report-tab-container > div:nth-of-type(2) > div > div:nth-of-type(2) a").ClickAsync();
                await ScreenShotLog(page, 5000);

                Log("Sending data parameters...");
                await IsElementExist("input[name=\"asOfDate\"]", page);
                var element = await page.QuerySelectorAsync("input[name=\"asOfDate\"]");
                await element.FillAsync(GetLastDayOfPrevMonth() + GetPrevMonth() + GetPrevYear());
                await ScreenShotLog(page);

                await IsElementExist("div:nth-of-type(9) div.text-right > button", page);
                await page.Locator("div:nth-of-type(9) div.text-right > button").ClickAsync();
                await ScreenShotLog(page, 10000);
                Log("Try to click the 'Download' button and waiting for download...");
                {
                    await IsElementExist("div.tab-control button.dropdown-toggle", page);
                    await page.Locator("div.tab-control button.dropdown-toggle").ClickAsync();
                    await page.WaitForTimeoutAsync(5000);
                    await IsElementExist("#module-accurate__report__report li:nth-of-type(2) > a", page);
                    var waitForDownloadTask = page.WaitForDownloadAsync();
                    await page.Locator("#module-accurate__report__report li:nth-of-type(2) > a").ClickAsync();

                    var download = await waitForDownloadTask;
                    await download.SaveAsAsync($"{strAppDownloadFolder}/Master_Outlet.xlsx");
                    Log($"Downloaded file path: {await download.PathAsync()}");
                }
            }
            await LogoutFunction(page);
            //Flag it as succesfull simulation
            return allpageprocessSuccess = true;
        }
        catch (Exception ex)
        {
            Log($"Error during web automation: {ex.Message}");
            return allpageprocessSuccess;
        }
        finally
        {

        }
    }

    private static async Task LogoutFunction(IPage page)
    {
        Log("Logging out...");
        Log("");
        var element = await page.QuerySelectorAsync("#dropdown-user--0");
        await element.ClickAsync();
        await page.WaitForTimeoutAsync(2000);
        await page.ScreenshotAsync(new() { Path = $"{strLogsFolder}/SS-No99.png" });
        element = await page.QuerySelectorAsync("//*[@id=\"accurate__init--0\"]/nav/div/div/div[2]/ul/li/a");
        await element.ClickAsync();
    }

    static async Task<bool> IsElementExist(String NameIDLocator, IPage page)
    {
        var elementExists = await page.Locator(NameIDLocator).CountAsync() > 0;
        var elementVisible = await page.Locator(NameIDLocator).IsVisibleAsync();
        if (enableElementCheckLogging == "Y")
        {
            Log($"Element exists: {elementExists} and is visible: {elementVisible}");
        }
        return elementExists;
        //return elementVisible;
    }

    private static int picCounter = 1;
    static async Task ScreenShotLog( IPage page, int delay = 1000)
    {
        await page.WaitForTimeoutAsync(delay);
        await page.ScreenshotAsync(new() { Path = $"{strLogsFolder}/SS-No" + picCounter + ".png" });
        picCounter += 1;
    }

    // Step 3: Zip and send files
    static void ZipAndSendFile()
    {
        try
        {
            Log("Checking and deleting existing ZIP files...");
            CheckAndDeleteZipFile(strDataSharingFolder);
            var strDsPeriod = GetPrevYear() + GetPrevMonth();

            Log("Moving standart excel reports file to uploaded folder...");
            // move excels files to Datafolder
            var path = strAppDownloadFolder + @"\Master_Outlet.xlsx";
            var path2 = strUploadFolder + @"\ds-" + strDTid + "-" + strDTname + "-" + strDsPeriod + "_OUTLET.xlsx";
            File.Move(path, path2, true);
            path = strAppDownloadFolder + @"\Sales_Data.xlsx";
            path2 = strUploadFolder + @"\ds-" + strDTid + "-" + strDTname + "-" + strDsPeriod + "_SALES.xlsx";
            File.Move(path, path2, true);
            path = strAppDownloadFolder + @"\Repayment_Data.xlsx";
            path2 = strUploadFolder + @"\ds-" + strDTid + "-" + strDTname + "-" + strDsPeriod + "_AR.xlsx";
            File.Move(path, path2, true);

            // set zipping name for files
            Log("Zipping Transaction file(s)");
            var strZipFile = strDTid + "-" + strDTname + "_" + strDsPeriod + ".zip";
            ZipFile.CreateFromDirectory(strUploadFolder, strDataSharingFolder + Path.DirectorySeparatorChar + strZipFile);

            // Send the ZIP file to the API server 
            Log("Sending ZIP file to the API server...");
            var strStatusCode = "0"; // varible for debugging Curl test
            strStatusCode = SendReq(strDataSharingFolder + Path.DirectorySeparatorChar + strZipFile, strSandBoxBool, "Y");
            Thread.Sleep(5000);
            if (strStatusCode == "200")
            {
                Log("DATA TRANSACTION SHARING - SELESAI");
            }
            else
            {
                Log("DATA TRANSACTION SHARING - ERROR, cUrl STATUS CODE :" + strStatusCode);
            }

            if (strAcctDataIncl=="Y")
            {
                //setup financial folder
                var financeFolder = strUploadFolder +  @"\financial report";
                if (Directory.Exists(financeFolder))
                {
                    Directory.Delete(financeFolder, true);
                    Directory.CreateDirectory(financeFolder);
                } 
                else
                {
                    Directory.CreateDirectory(financeFolder);
                }

                // move excels files to Datafolder
                Log("Moving Accounting excel reports file to uploaded folder...");
                path = strAppDownloadFolder + @"\Arus_Kas.xlsx";
                path2 = financeFolder + @"\ds-" + strDTid + "-" + strDTname + "-" + strDsPeriod + "_Arus_Kas.xlsx";
                File.Move(path, path2, true);
                path = strAppDownloadFolder + @"\Neraca_Saldo.xlsx";
                path2 = financeFolder + @"\ds-" + strDTid + "-" + strDTname + "-" + strDsPeriod + "_Neraca_Saldo.xlsx";
                File.Move(path, path2, true);
                path = strAppDownloadFolder + @"\Laporan_Laba_Rugi.xlsx";
                path2 = financeFolder + @"\ds-" + strDTid + "-" + strDTname + "-" + strDsPeriod + "_Laporan_Laba_Rugi.xlsx";
                File.Move(path, path2, true);

                // set zipping name for files
                Log("Zipping Financial file(s)");
                strZipFile = strDTid + "-" + strDTname + "-Financial Statement-" + GetDSPeriod() + ".zip";
                ZipFile.CreateFromDirectory(financeFolder, strDataSharingFolder + Path.DirectorySeparatorChar + strZipFile);

                // Send the ZIP file to the API server 
                Log("Sending ZIP file to the API server...");
                strStatusCode = SendReq(strDataSharingFolder + Path.DirectorySeparatorChar + strZipFile, strSandBoxBool, "Y");
                Thread.Sleep(5000);
                if (strStatusCode == "200")
                {
                    Log("DATA FINANCIAL SHARING - SELESAI");
                }
                else
                {
                    Log("DATA FINANCIAL SHARING - ERROR, cUrl STATUS CODE :" + strStatusCode);
                }
                //setup inventory folder
                var inventoryFolder = strUploadFolder + @"\inventory";
                if (Directory.Exists(inventoryFolder))
                {
                    Directory.Delete(inventoryFolder, true);
                    Directory.CreateDirectory(inventoryFolder);
                }
                else
                {
                    Directory.CreateDirectory(inventoryFolder);
                }
                Log("Moving inventory excel report file to uploaded folder...");
                path = strAppDownloadFolder + @"\Inventory.xlsx";
                path2 = inventoryFolder + @"\ds-" + strDTid + "-" + strDTname + "-" + strDsPeriod + "_Inventory.xlsx";
                File.Move(path, path2, true);
                // set zipping name for files
                Log("Zipping Inventory file(s)");
                strZipFile = strDTid + "-" + strDTname + "-Inventory-" + GetDSPeriod() + ".zip";
                ZipFile.CreateFromDirectory(inventoryFolder, strDataSharingFolder + Path.DirectorySeparatorChar + strZipFile);
                // Send the ZIP file to the API server 
                Log("Sending ZIP file to the API server...");
                strStatusCode = SendReq(strDataSharingFolder + Path.DirectorySeparatorChar + strZipFile, strSandBoxBool, "Y");
                Thread.Sleep(5000);
                if (strStatusCode == "200")
                {
                    Log("DATA INVENTORY SHARING - SELESAI");
                }
                else
                {
                    Log("DATA INVENTORY SHARING - ERROR, cUrl STATUS CODE :" + strStatusCode);
                }
            }
            // Send Log file to the API server 
            Log("Sending log file to the API server...");
            EndLogger();
            strStatusCode = SendReq(logFilePathAndName, strSandBoxBool, "Y");
            Thread.Sleep(5000);
            InitializeLogger(enableConsoleLogging);
        }
        catch (Exception ex)
        {
            // Log any exceptions that occur during file operations
            Log($"Error during ZIP and send: {ex.Message}");
        }
    }

    static void CheckAndDeleteZipFile(string folder)
    {
        try
        {

            var pattern = "*.zip";
            var zipFiles = Directory.EnumerateFiles(folder, pattern);
            foreach (var zipFile in zipFiles)
            {
                File.Delete(zipFile);
                Log($"Deleted ZIP file(s): {zipFile}");
            }

        }
        catch (Exception ex)
        {
            // Log any exceptions that occur during ZIP file deletion
            Log($"Error during ZIP file deletion: {ex.Message}");
        }
    }

    private static string SendReq(string strFileDataInfo, string strSandboxBool, string strSecureHTTP)
    {
        try
        {
            string text = "";
            string text2 = "";
            if (strSandboxBool == "Y")
            {
                text2 = "KQtbMk32csiJvm8XDAx2KnRAdbtP3YVAnJpF8R5cb2bcBr8boT3dTvGc23c6fqk2NknbxpdarsdF3M4V";
                text = ((!(strSecureHTTP == "Y")) ? "http://sandbox.fairbanc.app/api/documents" : "https://sandbox.fairbanc.app/api/documents");
            }
            else
            {
                text2 = "2S0VtpYzETxDrL6WClmxXXnOcCkNbR5nUCCLak6EHmbPbSSsJiTFTPNZrXKk2S0VtpYzETxDrL6WClmx";
                text = ((!(strSecureHTTP == "Y")) ? "http://dashboard.fairbanc.app/api/documents" : "https://dashboard.fairbanc.app/api/documents");
            }

            Log("Preparing to send a request to the API server...");
            HttpClient httpClient = new HttpClient();
            HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, text);
            MultipartFormDataContent multipartFormDataContent = new MultipartFormDataContent();
            multipartFormDataContent.Add(new StringContent(text2), "api_token");
            multipartFormDataContent.Add(new ByteArrayContent(File.ReadAllBytes(strFileDataInfo)), "file", Path.GetFileName(strFileDataInfo));
            httpRequestMessage.Content = multipartFormDataContent;
            HttpResponseMessage httpResponseMessage = httpClient.Send(httpRequestMessage);
            Thread.Sleep(5000);
            httpResponseMessage.EnsureSuccessStatusCode();
            var strResponseBody = httpResponseMessage.ToString();
            string[] array = strResponseBody.Split(':', ',');
            Log($"Response from API server: {array[1].Trim()}");
            return array[1].Trim();
        }
        catch (Exception ex)
        {
            // Log any exceptions that occur during the API request
            Log($"Error during API request: {ex.Message}");
            return "-1";
        }
    }

    static void DeleteAllSupportingFiles()
    {
        try
        {
            if (!Directory.Exists(strDataSharingFolder))
            {
                Directory.CreateDirectory(strDataSharingFolder);
            }else
            {
                Directory.Delete(strDataSharingFolder, true);
                Directory.CreateDirectory(strDataSharingFolder);
            }

            if (!Directory.Exists(strUploadFolder))
            {
                Directory.CreateDirectory(strUploadFolder);
            }else
            {
                Directory.Delete(strUploadFolder,true);
                Directory.CreateDirectory(strUploadFolder);
            }

            Log("Deleting supporting Excel files and image files...");

            // Delete Excel files
            var supportFiles1 = Directory.EnumerateFiles(strAppDownloadFolder, "*.xl*");
            foreach (var excelFile in supportFiles1)
            {
                File.Delete(excelFile);
                Log($"Deleted Excel file: {excelFile}");
            }

            // Delete image files
            var supportFiles2 = Directory.EnumerateFiles(strLogsFolder, "SS-*.png");
            foreach (var pictureFiles in supportFiles2)
            {
                File.Delete(pictureFiles);
                Log($"Deleted image file: {pictureFiles}");
            }

            // Delete log files
            var supportFiles3 = Directory.EnumerateFiles(strAppDownloadFolder, "Debug-*.log");
            foreach (var LogFiles in supportFiles3)
            {
                try
                {
                    File.Delete(LogFiles);
                    Log($"Deleted log file: {LogFiles}");
                }
                finally 
                {   // do nothing
                }
            }

            Log("Deleted all supporting files.");
        }
        catch (Exception ex)
        {
            // Log any exceptions that occur during file deletion
            Log($"Error during file deletion: {ex.Message}");
        }
    }

    private static void EndLogger()
    {
        // Flush and close the listener
        logListener.Close();

        // Remove the listener from the trace sources
        Trace.Listeners.Remove(logListener);
    }

    static void Log(string message)
    {
        Trace.WriteLine(message);
    }
}