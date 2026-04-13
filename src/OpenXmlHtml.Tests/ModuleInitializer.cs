public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        VerifierSettings.DontScrubDateTimes();
        VerifierSettings.DontScrubGuids();
        VerifyImageMagick.RegisterComparers(.5);
        VerifierSettings.InitializePlugins();
        VerifierSettings.UniqueForRuntime();
        VerifyOpenXmlConverter.Initialize();
    }
}
