#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;   // �f�o�b�O�Ŏg�p
using System.IO;  // �t�H���_���Q�Ƃ��邽�߂Ɏg�p
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
        /// <summary> �����l�̐錾
        /// </summary>
        const string _introLabName = "IntroCs";
        const string _uiLabName = "UiCs";
        const string _dllExtension = ".dll";
        // �摜�̂���T�u�f�B���N�g���̖��O
        const string _imageFolderName = "Ui_Image";
        // �R�}���h���`�����}�l�[�W�hdll�̏ꏊ
        string _introLabPath;
        // �A�C�R���p�摜�̈ʒu
        string _imageFolder;


        public Result OnShutdown(UIControlledApplication app)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication app)
        {
            // �O���A�v���P�[�V�����̃f�B���N�g��:
            string dir = Path.GetDirectoryName(
                System.Reflection.Assembly
                .GetExecutingAssembly().Location );

            // �O���R�}���h�p�X:
            _introLabPath = Path.Combine( dir, _introLabName + _dllExtension);

            if (!File.Exists(_introLabPath))
            {
                TaskDialog.Show("UIRibbon", "External command assembly not found: " + _introLabPath);
                return Result.Failed;
            }

            // �C���[�W�p�X:
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

            // �p�l���փ{�^���̒ǉ�
            PushButton pushButtonHello = 
                panel.AddItem(pushButtonDataHello ) as PushButton;

            // �A�C�R���̒ǉ�
            pushButtonHello.LargeImage = NewBitmapImage("ImgHelloWorld.png");

            // �c�[���`�b�v�̒ǉ�
            pushButtonHello.ToolTip = "�V���v���ȃv�b�V���{�^��";

        }

        /// <summary>
        /// AddPushButton02
        /// </summary>
        public void AddPushButton02(RibbonPanel panel)
        {
            PushButtonData pushButtonDataHello02
                = new PushButtonData("PushButtonHello02", "Hello World",
                                      _introLabPath, _introLabName + ".HelloWorldSimple");

            // �p�l���փ{�^���̒ǉ�
            PushButton pushButtonHello02 =
                panel.AddItem(pushButtonDataHello02) as PushButton;

            // �A�C�R���̒ǉ�
            pushButtonHello02.LargeImage = NewBitmapImage("ImgHelloWorld.png");

            // �c�[���`�b�v�̒ǉ�
            pushButtonHello02.ToolTip = "�V���v���ȃv�b�V���{�^��";

        }

        /// <summary> 
        /// AddSplitButton
        /// </summary>
        public void AddSplitButton (RibbonPanel panel) 
        {
            // 3�̃v�b�V���{�^�����X�v���b�g�{�^���ɓ����
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

            // �X�v���b�g�{�^���̍쐬
            SplitButtonData splitBtnData =
                new SplitButtonData("SplitButton", "Split Button");

            SplitButton splitBtn = panel.AddItem(splitBtnData ) as SplitButton;
            splitBtn.AddPushButton(pushButtonData1);
            splitBtn.AddPushButton(pushButtonData2);
            splitBtn.AddPushButton(pushButtonData3);
        }


        /// <summary> 
        /// ���[�e�B���e�B�֐�
        /// </summary>

        /// <summary>
        /// ����e�f�B���N�g���ɂ���A
        /// �^����ꂽ�^�[�Q�b�g�������T�u�f�B���N�g����������֐�
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
        /// �C���[�W�t�H���_����V�����A�C�R���̃r�b�g�}�b�v��ǂݍ��ފ֐�
        /// </summary>
        BitmapImage NewBitmapImage (string imageName)
        {
            return new BitmapImage(new Uri(Path.Combine(_imageFolder, imageName)));
        }
    }
}
