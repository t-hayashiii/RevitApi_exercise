#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;   // デバッグで使用
using System.IO;  // フォルダを参照するために使用
using System.Windows.Media.Imaging;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
// using Util;
#endregion

namespace UiCs
{
    public class UIRibbon : IExternalApplication
    {
        /// <summary> 初期値の宣言
        /// </summary>
        const string _introLabName = "IntroCs";
        const string _uiLabName = "UiCs";
        const string _dllExtension = ".dll";
        // 画像のあるサブディレクトリの名前
        const string _imageFolderName = "Ui_Image";
        // コマンドを定義したマネージドdllの場所
        string _introLabPath;
        // アイコン用画像の位置
        string _imageFolder;


        public Result OnShutdown(UIControlledApplication app)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication app)
        {
            // 外部アプリケーションのディレクトリ:
            string dir = Path.GetDirectoryName(
                System.Reflection.Assembly
                .GetExecutingAssembly().Location );

            // 外部コマンドパス:
            _introLabPath = Path.Combine( dir, _introLabName + _dllExtension);

            if (!File.Exists(_introLabPath))
            {
                TaskDialog.Show("UIRibbon", "External command assembly not found: " + _introLabPath);
                return Result.Failed;
            }

            // イメージパス:
            _imageFolder = FindFolderInParents(dir, _imageFolderName);

            if (null == _imageFolder
              || !Directory.Exists(_imageFolder))
            {
                TaskDialog.Show(
                  "UIRibbon",
                  string.Format(
                    "No image folder named '{0}' found in the parent directories of '{1}.",
                    _imageFolderName, dir));

                return Result.Failed;
            }


            AddRibbonSampler(app);

            return Result.Succeeded;
        }


        /// <summary>
        /// AddRibbonSampler
        /// </summary>
        public void AddRibbonSampler (UIControlledApplication app)
        {
            app.CreateRibbonTab("Ribbon Sampler");
            RibbonPanel panel =
                app.CreateRibbonPanel("Ribbon Sampler", "Ribbon Sampler");

            // (2.1) add a simple push button for Hello World 
            AddPushButton(panel);

            // (2.2) add split buttons for "Command Data", "DB Element" and "Element Filtering" 
            AddSplitButton(panel);

            // (2.3) add pulldown buttons for "Command Data", "DB Element" and "Element Filtering"

            // (2.4) add radio/toggle buttons for "Command Data", "DB Element" and "Element Filtering" 
            // we put it on the slide-out below. 
            //AddRadioButton(panel);
            //panel.AddSeparator();

            // (2.5) add text box - TBD: this is used with the conjunction with event. Probably too complex for day one training. 
            //  for now, without event. 
            // we put it on the slide-out below. 
            //AddTextBox(panel);
            //panel.AddSeparator();

            // (2.6) combo box - TBD: this is used with the conjunction with event. Probably too complex for day one training. 
            // For now, without event. show two groups: Element Bascis (3 push buttons) and Modification/Creation (2 push button)  


            // (2.7) stacked items - 1. hello world push button, 2. pulldown element bscis (command data, DB element, element filtering) 
            // 3. pulldown modification/creation(element modification, model creation). 


            // (2.8) slide out - if you don't have enough space, you can add additional space below the panel. 
            // anything which comes after this will be on the slide out. 
            panel.AddSlideOut();

            AddPushButton02(panel);


            // (2.4) radio button - what it is 


            // (2.5) text box - what it is 



        }

        /// <summary>
        /// AddPushButton
        /// </summary>
        public void AddPushButton(RibbonPanel panel)
        {
            PushButtonData pushButtonDataHello
                = new PushButtonData("PushButtonHello", "Hello World",
                                      _introLabPath, _introLabName + ".HelloWorldSimple");

            // パネルへボタンの追加
            PushButton pushButtonHello = 
                panel.AddItem(pushButtonDataHello ) as PushButton;

            // アイコンの追加
            pushButtonHello.LargeImage = NewBitmapImage("ImgHelloWorld.png");

            // ツールチップの追加
            pushButtonHello.ToolTip = "シンプルなプッシュボタン";

        }

        /// <summary>
        /// AddPushButton02
        /// </summary>
        public void AddPushButton02(RibbonPanel panel)
        {
            PushButtonData pushButtonDataHello02
                = new PushButtonData("PushButtonHello02", "Hello World",
                                      _introLabPath, _introLabName + ".HelloWorldSimple");

            // パネルへボタンの追加
            PushButton pushButtonHello02 =
                panel.AddItem(pushButtonDataHello02) as PushButton;

            // アイコンの追加
            pushButtonHello02.LargeImage = NewBitmapImage("ImgHelloWorld.png");

            // ツールチップの追加
            pushButtonHello02.ToolTip = "シンプルなプッシュボタン";

        }

        /// <summary> 
        /// AddSplitButton
        /// </summary>
        public void AddSplitButton (RibbonPanel panel) 
        {
            // 3つのプッシュボタンをスプリットボタンに入れる
            // #1
            PushButtonData pushButtonData1 =
                new PushButtonData("SplitCommandData", "Command Data", 
                    _introLabPath, _introLabName + ".CommandData");
            pushButtonData1.LargeImage = NewBitmapImage("ImgHelloWorld.png");

            // #2
            PushButtonData pushButtonData2 =
                new PushButtonData("SplitDbElement", "DB Element",
                    _introLabPath, _introLabName + ".DBElement");
            pushButtonData2.LargeImage = NewBitmapImage("ImgHelloWorld.png");

            // #3
            PushButtonData pushButtonData3 =
                new PushButtonData("SplitElementFiltering", "ElementFiltering",
                    _introLabPath, _introLabName + ".ElementFiltering");
            pushButtonData3.LargeImage = NewBitmapImage("ImgHelloWorld.png");

            // スプリットボタンの作成
            SplitButtonData splitBtnData =
                new SplitButtonData("SplitButton", "Split Button");

            SplitButton splitBtn = panel.AddItem(splitBtnData ) as SplitButton;
            splitBtn.AddPushButton(pushButtonData1);
            splitBtn.AddPushButton(pushButtonData2);
            splitBtn.AddPushButton(pushButtonData3);
        }


        /// <summary> 
        /// ユーティリティ関数
        /// </summary>

        /// <summary>
        /// ある親ディレクトリにある、
        /// 与えられたターゲット名を持つサブディレクトリ検索する関数
        /// </summary>
        string FindFolderInParents(string path, string target)
        {
            Debug.Assert(Directory.Exists(path),
                "expected an existing directory to start search in ");

            string s;

            do
            {
                s = Path.Combine(path, target);
                if (Directory.Exists(s))
                {
                    return s;
                }
                path = Path.GetDirectoryName(path);
            }
            while (null != path);

            return null;
        }

        /// <summary>
        /// イメージフォルダから新しいアイコンのビットマップを読み込む関数
        /// </summary>
        BitmapImage NewBitmapImage (string imageName)
        {
            return new BitmapImage(new Uri(Path.Combine(_imageFolder, imageName)));
        }
    }
}
