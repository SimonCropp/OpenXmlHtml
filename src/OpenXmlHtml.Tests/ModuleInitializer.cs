public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        VerifierSettings.DontScrubDateTimes();
        VerifierSettings.DontScrubGuids();
        VerifierSettings.UseSsimForPng();
        VerifierSettings.InitializePlugins();
        VerifierSettings.UniqueForRuntime();
        VerifyOpenXmlConverter.Initialize();
    }
}
