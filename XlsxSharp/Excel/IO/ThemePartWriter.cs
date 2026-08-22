using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.IO;

internal class ThemePartWriter
{
    internal static void GenerateContent(ThemePart themePart, XLTheme theme)
    {
        Theme theme1 = new() { Name = "Office Theme" };
        theme1.AddNamespaceDeclaration(
            "a",
            "http://schemas.openxmlformats.org/drawingml/2006/main"
        );

        ThemeElements themeElements1 = new();

        ColorScheme colorScheme1 = new() { Name = "Office" };

        Dark1Color dark1Color1 = new();
        SystemColor systemColor1 = new()
        {
            Val = SystemColorValues.WindowText,
            LastColor = theme.Text1.Color.ToHex().Substring(2),
        };

        dark1Color1.AppendChild(systemColor1);

        Light1Color light1Color1 = new();
        SystemColor systemColor2 = new()
        {
            Val = SystemColorValues.Window,
            LastColor = theme.Background1.Color.ToHex().Substring(2),
        };

        light1Color1.AppendChild(systemColor2);

        Dark2Color dark2Color1 = new();
        RgbColorModelHex rgbColorModelHex1 = new() { Val = theme.Text2.Color.ToHex().Substring(2) };

        dark2Color1.AppendChild(rgbColorModelHex1);

        Light2Color light2Color1 = new();
        RgbColorModelHex rgbColorModelHex2 = new()
        {
            Val = theme.Background2.Color.ToHex().Substring(2),
        };

        light2Color1.AppendChild(rgbColorModelHex2);

        Accent1Color accent1Color1 = new();
        RgbColorModelHex rgbColorModelHex3 = new()
        {
            Val = theme.Accent1.Color.ToHex().Substring(2),
        };

        accent1Color1.AppendChild(rgbColorModelHex3);

        Accent2Color accent2Color1 = new();
        RgbColorModelHex rgbColorModelHex4 = new()
        {
            Val = theme.Accent2.Color.ToHex().Substring(2),
        };

        accent2Color1.AppendChild(rgbColorModelHex4);

        Accent3Color accent3Color1 = new();
        RgbColorModelHex rgbColorModelHex5 = new()
        {
            Val = theme.Accent3.Color.ToHex().Substring(2),
        };

        accent3Color1.AppendChild(rgbColorModelHex5);

        Accent4Color accent4Color1 = new();
        RgbColorModelHex rgbColorModelHex6 = new()
        {
            Val = theme.Accent4.Color.ToHex().Substring(2),
        };

        accent4Color1.AppendChild(rgbColorModelHex6);

        Accent5Color accent5Color1 = new();
        RgbColorModelHex rgbColorModelHex7 = new()
        {
            Val = theme.Accent5.Color.ToHex().Substring(2),
        };

        accent5Color1.AppendChild(rgbColorModelHex7);

        Accent6Color accent6Color1 = new();
        RgbColorModelHex rgbColorModelHex8 = new()
        {
            Val = theme.Accent6.Color.ToHex().Substring(2),
        };

        accent6Color1.AppendChild(rgbColorModelHex8);

        Hyperlink hyperlink1 = new();
        RgbColorModelHex rgbColorModelHex9 = new()
        {
            Val = theme.Hyperlink.Color.ToHex().Substring(2),
        };

        hyperlink1.AppendChild(rgbColorModelHex9);

        FollowedHyperlinkColor followedHyperlinkColor1 = new();
        RgbColorModelHex rgbColorModelHex10 = new()
        {
            Val = theme.FollowedHyperlink.Color.ToHex().Substring(2),
        };

        followedHyperlinkColor1.AppendChild(rgbColorModelHex10);

        colorScheme1.AppendChild(dark1Color1);
        colorScheme1.AppendChild(light1Color1);
        colorScheme1.AppendChild(dark2Color1);
        colorScheme1.AppendChild(light2Color1);
        colorScheme1.AppendChild(accent1Color1);
        colorScheme1.AppendChild(accent2Color1);
        colorScheme1.AppendChild(accent3Color1);
        colorScheme1.AppendChild(accent4Color1);
        colorScheme1.AppendChild(accent5Color1);
        colorScheme1.AppendChild(accent6Color1);
        colorScheme1.AppendChild(hyperlink1);
        colorScheme1.AppendChild(followedHyperlinkColor1);

        FontScheme fontScheme2 = new() { Name = "Office" };

        MajorFont majorFont1 = new();
        LatinFont latinFont1 = new() { Typeface = "Cambria" };
        EastAsianFont eastAsianFont1 = new() { Typeface = "" };
        ComplexScriptFont complexScriptFont1 = new() { Typeface = "" };
        SupplementalFont supplementalFont1 = new()
        {
            Script = "Jpan",
            Typeface = "ＭＳ Ｐゴシック",
        };
        SupplementalFont supplementalFont2 = new() { Script = "Hang", Typeface = "맑은 고딕" };
        SupplementalFont supplementalFont3 = new() { Script = "Hans", Typeface = "宋体" };
        SupplementalFont supplementalFont4 = new() { Script = "Hant", Typeface = "新細明體" };
        SupplementalFont supplementalFont5 = new()
        {
            Script = "Arab",
            Typeface = "Times New Roman",
        };
        SupplementalFont supplementalFont6 = new()
        {
            Script = "Hebr",
            Typeface = "Times New Roman",
        };
        SupplementalFont supplementalFont7 = new() { Script = "Thai", Typeface = "Tahoma" };
        SupplementalFont supplementalFont8 = new() { Script = "Ethi", Typeface = "Nyala" };
        SupplementalFont supplementalFont9 = new() { Script = "Beng", Typeface = "Vrinda" };
        SupplementalFont supplementalFont10 = new() { Script = "Gujr", Typeface = "Shruti" };
        SupplementalFont supplementalFont11 = new() { Script = "Khmr", Typeface = "MoolBoran" };
        SupplementalFont supplementalFont12 = new() { Script = "Knda", Typeface = "Tunga" };
        SupplementalFont supplementalFont13 = new() { Script = "Guru", Typeface = "Raavi" };
        SupplementalFont supplementalFont14 = new() { Script = "Cans", Typeface = "Euphemia" };
        SupplementalFont supplementalFont15 = new()
        {
            Script = "Cher",
            Typeface = "Plantagenet Cherokee",
        };
        SupplementalFont supplementalFont16 = new()
        {
            Script = "Yiii",
            Typeface = "Microsoft Yi Baiti",
        };
        SupplementalFont supplementalFont17 = new()
        {
            Script = "Tibt",
            Typeface = "Microsoft Himalaya",
        };
        SupplementalFont supplementalFont18 = new() { Script = "Thaa", Typeface = "MV Boli" };
        SupplementalFont supplementalFont19 = new() { Script = "Deva", Typeface = "Mangal" };
        SupplementalFont supplementalFont20 = new() { Script = "Telu", Typeface = "Gautami" };
        SupplementalFont supplementalFont21 = new() { Script = "Taml", Typeface = "Latha" };
        SupplementalFont supplementalFont22 = new()
        {
            Script = "Syrc",
            Typeface = "Estrangelo Edessa",
        };
        SupplementalFont supplementalFont23 = new() { Script = "Orya", Typeface = "Kalinga" };
        SupplementalFont supplementalFont24 = new() { Script = "Mlym", Typeface = "Kartika" };
        SupplementalFont supplementalFont25 = new() { Script = "Laoo", Typeface = "DokChampa" };
        SupplementalFont supplementalFont26 = new() { Script = "Sinh", Typeface = "Iskoola Pota" };
        SupplementalFont supplementalFont27 = new()
        {
            Script = "Mong",
            Typeface = "Mongolian Baiti",
        };
        SupplementalFont supplementalFont28 = new()
        {
            Script = "Viet",
            Typeface = "Times New Roman",
        };
        SupplementalFont supplementalFont29 = new()
        {
            Script = "Uigh",
            Typeface = "Microsoft Uighur",
        };

        majorFont1.AppendChild(latinFont1);
        majorFont1.AppendChild(eastAsianFont1);
        majorFont1.AppendChild(complexScriptFont1);
        majorFont1.AppendChild(supplementalFont1);
        majorFont1.AppendChild(supplementalFont2);
        majorFont1.AppendChild(supplementalFont3);
        majorFont1.AppendChild(supplementalFont4);
        majorFont1.AppendChild(supplementalFont5);
        majorFont1.AppendChild(supplementalFont6);
        majorFont1.AppendChild(supplementalFont7);
        majorFont1.AppendChild(supplementalFont8);
        majorFont1.AppendChild(supplementalFont9);
        majorFont1.AppendChild(supplementalFont10);
        majorFont1.AppendChild(supplementalFont11);
        majorFont1.AppendChild(supplementalFont12);
        majorFont1.AppendChild(supplementalFont13);
        majorFont1.AppendChild(supplementalFont14);
        majorFont1.AppendChild(supplementalFont15);
        majorFont1.AppendChild(supplementalFont16);
        majorFont1.AppendChild(supplementalFont17);
        majorFont1.AppendChild(supplementalFont18);
        majorFont1.AppendChild(supplementalFont19);
        majorFont1.AppendChild(supplementalFont20);
        majorFont1.AppendChild(supplementalFont21);
        majorFont1.AppendChild(supplementalFont22);
        majorFont1.AppendChild(supplementalFont23);
        majorFont1.AppendChild(supplementalFont24);
        majorFont1.AppendChild(supplementalFont25);
        majorFont1.AppendChild(supplementalFont26);
        majorFont1.AppendChild(supplementalFont27);
        majorFont1.AppendChild(supplementalFont28);
        majorFont1.AppendChild(supplementalFont29);

        MinorFont minorFont1 = new();
        LatinFont latinFont2 = new() { Typeface = "Calibri" };
        EastAsianFont eastAsianFont2 = new() { Typeface = "" };
        ComplexScriptFont complexScriptFont2 = new() { Typeface = "" };
        SupplementalFont supplementalFont30 = new()
        {
            Script = "Jpan",
            Typeface = "ＭＳ Ｐゴシック",
        };
        SupplementalFont supplementalFont31 = new() { Script = "Hang", Typeface = "맑은 고딕" };
        SupplementalFont supplementalFont32 = new() { Script = "Hans", Typeface = "宋体" };
        SupplementalFont supplementalFont33 = new() { Script = "Hant", Typeface = "新細明體" };
        SupplementalFont supplementalFont34 = new() { Script = "Arab", Typeface = "Arial" };
        SupplementalFont supplementalFont35 = new() { Script = "Hebr", Typeface = "Arial" };
        SupplementalFont supplementalFont36 = new() { Script = "Thai", Typeface = "Tahoma" };
        SupplementalFont supplementalFont37 = new() { Script = "Ethi", Typeface = "Nyala" };
        SupplementalFont supplementalFont38 = new() { Script = "Beng", Typeface = "Vrinda" };
        SupplementalFont supplementalFont39 = new() { Script = "Gujr", Typeface = "Shruti" };
        SupplementalFont supplementalFont40 = new() { Script = "Khmr", Typeface = "DaunPenh" };
        SupplementalFont supplementalFont41 = new() { Script = "Knda", Typeface = "Tunga" };
        SupplementalFont supplementalFont42 = new() { Script = "Guru", Typeface = "Raavi" };
        SupplementalFont supplementalFont43 = new() { Script = "Cans", Typeface = "Euphemia" };
        SupplementalFont supplementalFont44 = new()
        {
            Script = "Cher",
            Typeface = "Plantagenet Cherokee",
        };
        SupplementalFont supplementalFont45 = new()
        {
            Script = "Yiii",
            Typeface = "Microsoft Yi Baiti",
        };
        SupplementalFont supplementalFont46 = new()
        {
            Script = "Tibt",
            Typeface = "Microsoft Himalaya",
        };
        SupplementalFont supplementalFont47 = new() { Script = "Thaa", Typeface = "MV Boli" };
        SupplementalFont supplementalFont48 = new() { Script = "Deva", Typeface = "Mangal" };
        SupplementalFont supplementalFont49 = new() { Script = "Telu", Typeface = "Gautami" };
        SupplementalFont supplementalFont50 = new() { Script = "Taml", Typeface = "Latha" };
        SupplementalFont supplementalFont51 = new()
        {
            Script = "Syrc",
            Typeface = "Estrangelo Edessa",
        };
        SupplementalFont supplementalFont52 = new() { Script = "Orya", Typeface = "Kalinga" };
        SupplementalFont supplementalFont53 = new() { Script = "Mlym", Typeface = "Kartika" };
        SupplementalFont supplementalFont54 = new() { Script = "Laoo", Typeface = "DokChampa" };
        SupplementalFont supplementalFont55 = new() { Script = "Sinh", Typeface = "Iskoola Pota" };
        SupplementalFont supplementalFont56 = new()
        {
            Script = "Mong",
            Typeface = "Mongolian Baiti",
        };
        SupplementalFont supplementalFont57 = new() { Script = "Viet", Typeface = "Arial" };
        SupplementalFont supplementalFont58 = new()
        {
            Script = "Uigh",
            Typeface = "Microsoft Uighur",
        };

        minorFont1.AppendChild(latinFont2);
        minorFont1.AppendChild(eastAsianFont2);
        minorFont1.AppendChild(complexScriptFont2);
        minorFont1.AppendChild(supplementalFont30);
        minorFont1.AppendChild(supplementalFont31);
        minorFont1.AppendChild(supplementalFont32);
        minorFont1.AppendChild(supplementalFont33);
        minorFont1.AppendChild(supplementalFont34);
        minorFont1.AppendChild(supplementalFont35);
        minorFont1.AppendChild(supplementalFont36);
        minorFont1.AppendChild(supplementalFont37);
        minorFont1.AppendChild(supplementalFont38);
        minorFont1.AppendChild(supplementalFont39);
        minorFont1.AppendChild(supplementalFont40);
        minorFont1.AppendChild(supplementalFont41);
        minorFont1.AppendChild(supplementalFont42);
        minorFont1.AppendChild(supplementalFont43);
        minorFont1.AppendChild(supplementalFont44);
        minorFont1.AppendChild(supplementalFont45);
        minorFont1.AppendChild(supplementalFont46);
        minorFont1.AppendChild(supplementalFont47);
        minorFont1.AppendChild(supplementalFont48);
        minorFont1.AppendChild(supplementalFont49);
        minorFont1.AppendChild(supplementalFont50);
        minorFont1.AppendChild(supplementalFont51);
        minorFont1.AppendChild(supplementalFont52);
        minorFont1.AppendChild(supplementalFont53);
        minorFont1.AppendChild(supplementalFont54);
        minorFont1.AppendChild(supplementalFont55);
        minorFont1.AppendChild(supplementalFont56);
        minorFont1.AppendChild(supplementalFont57);
        minorFont1.AppendChild(supplementalFont58);

        fontScheme2.AppendChild(majorFont1);
        fontScheme2.AppendChild(minorFont1);

        FormatScheme formatScheme1 = new() { Name = "Office" };

        FillStyleList fillStyleList1 = new();

        SolidFill solidFill1 = new();
        SchemeColor schemeColor1 = new() { Val = SchemeColorValues.PhColor };

        solidFill1.AppendChild(schemeColor1);

        GradientFill gradientFill1 = new() { RotateWithShape = true };

        GradientStopList gradientStopList1 = new();

        GradientStop gradientStop1 = new() { Position = 0 };

        SchemeColor schemeColor2 = new() { Val = SchemeColorValues.PhColor };
        Tint tint1 = new() { Val = 50000 };
        SaturationModulation saturationModulation1 = new() { Val = 300000 };

        schemeColor2.AppendChild(tint1);
        schemeColor2.AppendChild(saturationModulation1);

        gradientStop1.AppendChild(schemeColor2);

        GradientStop gradientStop2 = new() { Position = 35000 };

        SchemeColor schemeColor3 = new() { Val = SchemeColorValues.PhColor };
        Tint tint2 = new() { Val = 37000 };
        SaturationModulation saturationModulation2 = new() { Val = 300000 };

        schemeColor3.AppendChild(tint2);
        schemeColor3.AppendChild(saturationModulation2);

        gradientStop2.AppendChild(schemeColor3);

        GradientStop gradientStop3 = new() { Position = 100000 };

        SchemeColor schemeColor4 = new() { Val = SchemeColorValues.PhColor };
        Tint tint3 = new() { Val = 15000 };
        SaturationModulation saturationModulation3 = new() { Val = 350000 };

        schemeColor4.AppendChild(tint3);
        schemeColor4.AppendChild(saturationModulation3);

        gradientStop3.AppendChild(schemeColor4);

        gradientStopList1.AppendChild(gradientStop1);
        gradientStopList1.AppendChild(gradientStop2);
        gradientStopList1.AppendChild(gradientStop3);
        LinearGradientFill linearGradientFill1 = new() { Angle = 16200000, Scaled = true };

        gradientFill1.AppendChild(gradientStopList1);
        gradientFill1.AppendChild(linearGradientFill1);

        GradientFill gradientFill2 = new() { RotateWithShape = true };

        GradientStopList gradientStopList2 = new();

        GradientStop gradientStop4 = new() { Position = 0 };

        SchemeColor schemeColor5 = new() { Val = SchemeColorValues.PhColor };
        Shade shade1 = new() { Val = 51000 };
        SaturationModulation saturationModulation4 = new() { Val = 130000 };

        schemeColor5.AppendChild(shade1);
        schemeColor5.AppendChild(saturationModulation4);

        gradientStop4.AppendChild(schemeColor5);

        GradientStop gradientStop5 = new() { Position = 80000 };

        SchemeColor schemeColor6 = new() { Val = SchemeColorValues.PhColor };
        Shade shade2 = new() { Val = 93000 };
        SaturationModulation saturationModulation5 = new() { Val = 130000 };

        schemeColor6.AppendChild(shade2);
        schemeColor6.AppendChild(saturationModulation5);

        gradientStop5.AppendChild(schemeColor6);

        GradientStop gradientStop6 = new() { Position = 100000 };

        SchemeColor schemeColor7 = new() { Val = SchemeColorValues.PhColor };
        Shade shade3 = new() { Val = 94000 };
        SaturationModulation saturationModulation6 = new() { Val = 135000 };

        schemeColor7.AppendChild(shade3);
        schemeColor7.AppendChild(saturationModulation6);

        gradientStop6.AppendChild(schemeColor7);

        gradientStopList2.AppendChild(gradientStop4);
        gradientStopList2.AppendChild(gradientStop5);
        gradientStopList2.AppendChild(gradientStop6);
        LinearGradientFill linearGradientFill2 = new() { Angle = 16200000, Scaled = false };

        gradientFill2.AppendChild(gradientStopList2);
        gradientFill2.AppendChild(linearGradientFill2);

        fillStyleList1.AppendChild(solidFill1);
        fillStyleList1.AppendChild(gradientFill1);
        fillStyleList1.AppendChild(gradientFill2);

        LineStyleList lineStyleList1 = new();

        Outline outline1 = new()
        {
            Width = 9525,
            CapType = LineCapValues.Flat,
            CompoundLineType = CompoundLineValues.Single,
            Alignment = PenAlignmentValues.Center,
        };

        SolidFill solidFill2 = new();

        SchemeColor schemeColor8 = new() { Val = SchemeColorValues.PhColor };
        Shade shade4 = new() { Val = 95000 };
        SaturationModulation saturationModulation7 = new() { Val = 105000 };

        schemeColor8.AppendChild(shade4);
        schemeColor8.AppendChild(saturationModulation7);

        solidFill2.AppendChild(schemeColor8);
        PresetDash presetDash1 = new() { Val = PresetLineDashValues.Solid };

        outline1.AppendChild(solidFill2);
        outline1.AppendChild(presetDash1);

        Outline outline2 = new()
        {
            Width = 25400,
            CapType = LineCapValues.Flat,
            CompoundLineType = CompoundLineValues.Single,
            Alignment = PenAlignmentValues.Center,
        };

        SolidFill solidFill3 = new();
        SchemeColor schemeColor9 = new() { Val = SchemeColorValues.PhColor };

        solidFill3.AppendChild(schemeColor9);
        PresetDash presetDash2 = new() { Val = PresetLineDashValues.Solid };

        outline2.AppendChild(solidFill3);
        outline2.AppendChild(presetDash2);

        Outline outline3 = new()
        {
            Width = 38100,
            CapType = LineCapValues.Flat,
            CompoundLineType = CompoundLineValues.Single,
            Alignment = PenAlignmentValues.Center,
        };

        SolidFill solidFill4 = new();
        SchemeColor schemeColor10 = new() { Val = SchemeColorValues.PhColor };

        solidFill4.AppendChild(schemeColor10);
        PresetDash presetDash3 = new() { Val = PresetLineDashValues.Solid };

        outline3.AppendChild(solidFill4);
        outline3.AppendChild(presetDash3);

        lineStyleList1.AppendChild(outline1);
        lineStyleList1.AppendChild(outline2);
        lineStyleList1.AppendChild(outline3);

        EffectStyleList effectStyleList1 = new();

        EffectStyle effectStyle1 = new();

        EffectList effectList1 = new();

        OuterShadow outerShadow1 = new()
        {
            BlurRadius = 40000L,
            Distance = 20000L,
            Direction = 5400000,
            RotateWithShape = false,
        };

        RgbColorModelHex rgbColorModelHex11 = new() { Val = "000000" };
        Alpha alpha1 = new() { Val = 38000 };

        rgbColorModelHex11.AppendChild(alpha1);

        outerShadow1.AppendChild(rgbColorModelHex11);

        effectList1.AppendChild(outerShadow1);

        effectStyle1.AppendChild(effectList1);

        EffectStyle effectStyle2 = new();

        EffectList effectList2 = new();

        OuterShadow outerShadow2 = new()
        {
            BlurRadius = 40000L,
            Distance = 23000L,
            Direction = 5400000,
            RotateWithShape = false,
        };

        RgbColorModelHex rgbColorModelHex12 = new() { Val = "000000" };
        Alpha alpha2 = new() { Val = 35000 };

        rgbColorModelHex12.AppendChild(alpha2);

        outerShadow2.AppendChild(rgbColorModelHex12);

        effectList2.AppendChild(outerShadow2);

        effectStyle2.AppendChild(effectList2);

        EffectStyle effectStyle3 = new();

        EffectList effectList3 = new();

        OuterShadow outerShadow3 = new()
        {
            BlurRadius = 40000L,
            Distance = 23000L,
            Direction = 5400000,
            RotateWithShape = false,
        };

        RgbColorModelHex rgbColorModelHex13 = new() { Val = "000000" };
        Alpha alpha3 = new() { Val = 35000 };

        rgbColorModelHex13.AppendChild(alpha3);

        outerShadow3.AppendChild(rgbColorModelHex13);

        effectList3.AppendChild(outerShadow3);

        Scene3DType scene3DType1 = new();

        Camera camera1 = new() { Preset = PresetCameraValues.OrthographicFront };
        Rotation rotation1 = new()
        {
            Latitude = 0,
            Longitude = 0,
            Revolution = 0,
        };

        camera1.AppendChild(rotation1);

        LightRig lightRig1 = new()
        {
            Rig = LightRigValues.ThreePoints,
            Direction = LightRigDirectionValues.Top,
        };
        Rotation rotation2 = new()
        {
            Latitude = 0,
            Longitude = 0,
            Revolution = 1200000,
        };

        lightRig1.AppendChild(rotation2);

        scene3DType1.AppendChild(camera1);
        scene3DType1.AppendChild(lightRig1);

        Shape3DType shape3DType1 = new();
        BevelTop bevelTop1 = new() { Width = 63500L, Height = 25400L };

        shape3DType1.AppendChild(bevelTop1);

        effectStyle3.AppendChild(effectList3);
        effectStyle3.AppendChild(scene3DType1);
        effectStyle3.AppendChild(shape3DType1);

        effectStyleList1.AppendChild(effectStyle1);
        effectStyleList1.AppendChild(effectStyle2);
        effectStyleList1.AppendChild(effectStyle3);

        BackgroundFillStyleList backgroundFillStyleList1 = new();

        SolidFill solidFill5 = new();
        SchemeColor schemeColor11 = new() { Val = SchemeColorValues.PhColor };

        solidFill5.AppendChild(schemeColor11);

        GradientFill gradientFill3 = new() { RotateWithShape = true };

        GradientStopList gradientStopList3 = new();

        GradientStop gradientStop7 = new() { Position = 0 };

        SchemeColor schemeColor12 = new() { Val = SchemeColorValues.PhColor };
        Tint tint4 = new() { Val = 40000 };
        SaturationModulation saturationModulation8 = new() { Val = 350000 };

        schemeColor12.AppendChild(tint4);
        schemeColor12.AppendChild(saturationModulation8);

        gradientStop7.AppendChild(schemeColor12);

        GradientStop gradientStop8 = new() { Position = 40000 };

        SchemeColor schemeColor13 = new() { Val = SchemeColorValues.PhColor };
        Tint tint5 = new() { Val = 45000 };
        Shade shade5 = new() { Val = 99000 };
        SaturationModulation saturationModulation9 = new() { Val = 350000 };

        schemeColor13.AppendChild(tint5);
        schemeColor13.AppendChild(shade5);
        schemeColor13.AppendChild(saturationModulation9);

        gradientStop8.AppendChild(schemeColor13);

        GradientStop gradientStop9 = new() { Position = 100000 };

        SchemeColor schemeColor14 = new() { Val = SchemeColorValues.PhColor };
        Shade shade6 = new() { Val = 20000 };
        SaturationModulation saturationModulation10 = new() { Val = 255000 };

        schemeColor14.AppendChild(shade6);
        schemeColor14.AppendChild(saturationModulation10);

        gradientStop9.AppendChild(schemeColor14);

        gradientStopList3.AppendChild(gradientStop7);
        gradientStopList3.AppendChild(gradientStop8);
        gradientStopList3.AppendChild(gradientStop9);

        PathGradientFill pathGradientFill1 = new() { Path = PathShadeValues.Circle };
        FillToRectangle fillToRectangle1 = new()
        {
            Left = 50000,
            Top = -80000,
            Right = 50000,
            Bottom = 180000,
        };

        pathGradientFill1.AppendChild(fillToRectangle1);

        gradientFill3.AppendChild(gradientStopList3);
        gradientFill3.AppendChild(pathGradientFill1);

        GradientFill gradientFill4 = new() { RotateWithShape = true };

        GradientStopList gradientStopList4 = new();

        GradientStop gradientStop10 = new() { Position = 0 };

        SchemeColor schemeColor15 = new() { Val = SchemeColorValues.PhColor };
        Tint tint6 = new() { Val = 80000 };
        SaturationModulation saturationModulation11 = new() { Val = 300000 };

        schemeColor15.AppendChild(tint6);
        schemeColor15.AppendChild(saturationModulation11);

        gradientStop10.AppendChild(schemeColor15);

        GradientStop gradientStop11 = new() { Position = 100000 };

        SchemeColor schemeColor16 = new() { Val = SchemeColorValues.PhColor };
        Shade shade7 = new() { Val = 30000 };
        SaturationModulation saturationModulation12 = new() { Val = 200000 };

        schemeColor16.AppendChild(shade7);
        schemeColor16.AppendChild(saturationModulation12);

        gradientStop11.AppendChild(schemeColor16);

        gradientStopList4.AppendChild(gradientStop10);
        gradientStopList4.AppendChild(gradientStop11);

        PathGradientFill pathGradientFill2 = new() { Path = PathShadeValues.Circle };
        FillToRectangle fillToRectangle2 = new()
        {
            Left = 50000,
            Top = 50000,
            Right = 50000,
            Bottom = 50000,
        };

        pathGradientFill2.AppendChild(fillToRectangle2);

        gradientFill4.AppendChild(gradientStopList4);
        gradientFill4.AppendChild(pathGradientFill2);

        backgroundFillStyleList1.AppendChild(solidFill5);
        backgroundFillStyleList1.AppendChild(gradientFill3);
        backgroundFillStyleList1.AppendChild(gradientFill4);

        formatScheme1.AppendChild(fillStyleList1);
        formatScheme1.AppendChild(lineStyleList1);
        formatScheme1.AppendChild(effectStyleList1);
        formatScheme1.AppendChild(backgroundFillStyleList1);

        themeElements1.AppendChild(colorScheme1);
        themeElements1.AppendChild(fontScheme2);
        themeElements1.AppendChild(formatScheme1);
        ObjectDefaults objectDefaults1 = new();
        ExtraColorSchemeList extraColorSchemeList1 = new();

        theme1.AppendChild(themeElements1);
        theme1.AppendChild(objectDefaults1);
        theme1.AppendChild(extraColorSchemeList1);

        themePart.Theme = theme1;
    }
}
