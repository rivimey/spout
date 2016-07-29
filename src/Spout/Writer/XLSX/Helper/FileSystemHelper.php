<?php

namespace Box\Spout\Writer\XLSX\Helper;

use Box\Spout\Writer\XLSX\Internal\Worksheet;
use Rivimey\ZipStreamer\Deflate\COMPR;
use Rivimey\ZipStreamer\ZipStreamer;

const XML_EOL = "\n";

/**
 * Class FileSystemHelper
 * This class provides helper functions to help with the file system operations
 * like files/folders creation & deletion for XLSX files
 *
 * @package Box\Spout\Writer\XLSX\Helper
 */
class FileSystemHelper extends \Box\Spout\Common\Helper\FileSystemHelper
{
    const APP_NAME = 'Spout';

    const RELS_FOLDER_NAME = '_rels';
    const DOC_PROPS_FOLDER_NAME = 'docProps';
    const XL_FOLDER_NAME = 'xl';
    const WORKSHEETS_FOLDER_NAME = 'worksheets';

    const RELS_FILE_NAME = '.rels';
    const APP_XML_FILE_NAME = 'app.xml';
    const CORE_XML_FILE_NAME = 'core.xml';
    const CONTENT_TYPES_XML_FILE_NAME = '[Content_Types].xml';
    const WORKBOOK_XML_FILE_NAME = 'workbook.xml';
    const WORKBOOK_RELS_XML_FILE_NAME = 'workbook.xml.rels';
    const STYLES_XML_FILE_NAME = 'styles.xml';

    protected $ooXmlPackageNs = "http://schemas.openxmlformats.org/package/2006";
    protected $ooXmlOfficeDocNs = "http://schemas.openxmlformats.org/officeDocument/2006";
    protected $ooXmlSheetDocNs = "http://schemas.openxmlformats.org/spreadsheetml/2006";
    protected $ooCTypeDoc = "application/vnd.openxmlformats-officedocument";
    protected $ooCTypePkg = "application/vnd.openxmlformats-package";

    /** @var string Path to the root folder inside the temp folder where the files to create the XLSX will be stored */
    protected $rootFolder;

    /** @var string Path to the "_rels" folder inside the root folder */
    protected $relsFolder;

    /** @var string Path to the "docProps" folder inside the root folder */
    protected $docPropsFolder;

    /** @var string Path to the "xl" folder inside the root folder */
    protected $xlFolder;

    /** @var string Path to the "_rels" folder inside the "xl" folder */
    protected $xlRelsFolder;

    /** @var string Path to the "worksheets" folder inside the "xl" folder */
    protected $xlWorksheetsFolder;

    /** @var  ZipStreamer $zipStream */
    protected $zipStream;

    /**
     * Set the ZipStream object to be used to write files.
     *
     * @return ZipStreamer
     */
    public function setZipStream(ZipStreamer $zipStream)
    {
        $this->zipStream = $zipStream;
    }

    /**
     * Return the ZipStream object associated with this file helper.
     *
     * @return ZipStreamer
     */
    public function getZipStream()
    {
        return $this->zipStream;
    }

    /**
     * Creates an empty folder with the given name under the given parent folder.
     *
     * @param string $parentFolderPath The parent folder path under which the folder is going to be created
     * @param string $folderName The name of the folder to create
     * @return string Path of the created folder
     * @throws \Box\Spout\Common\Exception\IOException If unable to create the folder or if the folder path is not inside of the base folder
     */
    public function createFolder($parentFolderPath, $folderName)
    {
        if (!empty($parentFolderPath)) {
            $folderPath = $parentFolderPath . '/' . $folderName;
        } else {
            $folderPath = $folderName;
        }

        // TODO: Could explicitly add the subfolder but it isn't needed... Is there a reason to do so?
        // $this->zipStream->addEmptyDir($folderPath);
        return $folderPath;
    }

    /**
     * Creates a file with the given name and content in the given folder.
     * The parent folder must exist.
     *
     * @param string $parentFolderPath The parent folder path where the file is going to be created
     * @param string $fileName The name of the file to create
     * @param string $fileContents The contents of the file to create
     * @return string Path of the created file
     * @throws \Box\Spout\Common\Exception\IOException If unable to create the file or if the file path is not inside of the base folder
     */
    public function createFileWithContents($parentFolderPath, $fileName, $fileContents)
    {
        if (!empty($parentFolderPath)) {
            $filePath = $parentFolderPath . '/' . $fileName;
        } else {
            $filePath = $fileName;
        }
        $this->zipStream->addFileFromString($fileContents, $filePath);

        return $filePath;
    }

    /**
     * Not implemented: Delete the file
     *
     * @throws \Box\Spout\Common\Exception\WriterException
     */
    public function deleteFile($filePath)
    {
        throw new WriterException('Unable to delete file in streamed zip output');
    }

    /**
     * Not implemented: Delete the folder
     *
     * @throws \Box\Spout\Common\Exception\WriterException
     */
    public function deleteFolderRecursively($folderPath)
    {
        throw new WriterException('Unable to delete folder in streamed zip output');
    }

    /**
     * @return string
     */
    public function getRootFolder()
    {
        return $this->rootFolder;
    }

    /**
     * @return string
     */
    public function getXlFolder()
    {
        return $this->xlFolder;
    }

    /**
     * @return string
     */
    public function getXlWorksheetsFolder()
    {
        return $this->xlWorksheetsFolder;
    }

    /**
     * Creates all the folders needed to create a XLSX file, as well as the files that won't change.
     *
     * @return void
     * @throws \Box\Spout\Common\Exception\IOException If unable to create at least one of the base folders
     */
    public function createBaseFilesAndFolders()
    {
        $this
            ->createRootFolder()
            ->createRelsFolderAndFile()
            ->createDocPropsFolderAndFiles()
            ->createXlFolderAndSubFolders();
    }

    /**
     * Creates the folder that will be used as root
     *
     * @return FileSystemHelper
     * @throws \Box\Spout\Common\Exception\IOException If unable to create the folder
     */
    protected function createRootFolder()
    {
        $this->rootFolder = $this->createFolder('', $this->baseFolderPath);

        return $this;
    }

    /**
     * Creates the "_rels" folder under the root folder as well as the ".rels" file in it
     *
     * @return FileSystemHelper
     * @throws \Box\Spout\Common\Exception\IOException If unable to create the folder or the ".rels" file
     */
    protected function createRelsFolderAndFile()
    {
        $this->relsFolder = $this->createFolder($this->rootFolder, self::RELS_FOLDER_NAME);

        $this->createRelsFile();

        return $this;
    }

    /**
     * Creates the ".rels" file under the "_rels" folder (under root)
     *
     * @return FileSystemHelper
     * @throws \Box\Spout\Common\Exception\IOException If unable to create the file
     */
    protected function createRelsFile()
    {
        $relsFileContents = <<<EOD
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="{$this->ooXmlPackageNs}/relationships">
    <Relationship Id="rIdWbk" Type="{$this->ooXmlOfficeDocNs}/relationships/officeDocument" Target="xl/workbook.xml"/>
    <Relationship Id="rIdCore" Type="{$this->ooXmlPackageNs}/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rIdApp" Type="{$this->ooXmlOfficeDocNs}/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
EOD;

        $this->createFileWithContents($this->relsFolder, self::RELS_FILE_NAME, $relsFileContents);
        return $this;
    }

    /**
     * Creates the "docProps" folder under the root folder as well as the "app.xml" and "core.xml" files in it.
     *
     * @return FileSystemHelper
     * @throws \Box\Spout\Common\Exception\IOException If unable to create the folder or one of the files
     */
    protected function createDocPropsFolderAndFiles()
    {
        $this->docPropsFolder = $this->createFolder($this->rootFolder, self::DOC_PROPS_FOLDER_NAME);

        $this->createAppXmlFile();
        $this->createCoreXmlFile();

        return $this;
    }

    /**
     * Creates the "app.xml" file under the "docProps" folder
     *
     * @return FileSystemHelper
     * @throws \Box\Spout\Common\Exception\IOException If unable to create the file
     */
    protected function createAppXmlFile()
    {
        $appName = self::APP_NAME;
        // TODO: xmlns:vt="{$this->ooDocXmlns}/docPropsVTypes" is needed for TitlesOfParts or HeadingPairs.
        // Plausibly add: <Company> <AppVersion>
        // Maybe also: <ScaleCrop> <TitlesOfParts> <LinksUpToDate> <SharedDoc>
        $appXmlFileContents = <<<EOD
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="{$this->ooDocXmlns}/extended-properties">
  <Application>$appName</Application>
  <TotalTime>0</TotalTime>
</Properties>
EOD;

        $this->createFileWithContents($this->docPropsFolder, self::APP_XML_FILE_NAME, $appXmlFileContents);

        return $this;
    }

    /**
     * Creates the "core.xml" file under the "docProps" folder
     *
     * @return FileSystemHelper
     * @throws \Box\Spout\Common\Exception\IOException If unable to create the file
     */
    protected function createCoreXmlFile()
    {
        $createdDate = (new \DateTime())->format(\DateTime::W3C);
        $coreXmlFileContents = <<<EOD
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="{$this->ooXmlPackageNs}/metadata/core-properties" 
  xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" 
  xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dcterms:created xsi:type="dcterms:W3CDTF">$createdDate</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">$createdDate</dcterms:modified>
  <cp:revision>0</cp:revision>
</cp:coreProperties>
EOD;

        $this->createFileWithContents($this->docPropsFolder, self::CORE_XML_FILE_NAME, $coreXmlFileContents);
        return $this;
    }

    /**
     * Creates the "xl" folder under the root folder as well as its subfolders
     *
     * @return FileSystemHelper
     * @throws \Box\Spout\Common\Exception\IOException If unable to create at least one of the folders
     */
    protected function createXlFolderAndSubFolders()
    {
        $this->xlFolder = $this->createFolder($this->rootFolder, self::XL_FOLDER_NAME);
        $this->createXlRelsFolder();
        $this->createXlWorksheetsFolder();
        return $this;
    }

    /**
     * Creates the "_rels" folder under the "xl" folder
     *
     * @return FileSystemHelper
     * @throws \Box\Spout\Common\Exception\IOException If unable to create the folder
     */
    protected function createXlRelsFolder()
    {
        $this->xlRelsFolder = $this->createFolder($this->xlFolder, self::RELS_FOLDER_NAME);
        return $this;
    }

    /**
     * Creates the "worksheets" folder under the "xl" folder
     *
     * @return FileSystemHelper
     * @throws \Box\Spout\Common\Exception\IOException If unable to create the folder
     */
    protected function createXlWorksheetsFolder()
    {
        $this->xlWorksheetsFolder = $this->createFolder($this->xlFolder, self::WORKSHEETS_FOLDER_NAME);
        return $this;
    }

    /**
     * Creates the "[Content_Types].xml" file under the root folder
     *
     * @param Worksheet[] $worksheets
     * @return FileSystemHelper
     */
    public function createContentTypesFile($worksheets, $shouldUseInlineStrings)
    {
        $contentTypesXmlFileContents = <<<EOD
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="{$this->ooXmlPackageNs}/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="{$this->ooCTypePkg}.relationships+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="{$this->ooCTypeDoc}.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/_rels/.rels" ContentType="{$this->ooCTypePkg}.relationships+xml"/>

EOD;

        /** @var Worksheet $worksheet */
        foreach ($worksheets as $worksheet) {
            $contentTypesXmlFileContents .=
                '  <Override PartName="/xl/worksheets/sheet' . $worksheet->getId() . '.xml" ContentType="' . $this->ooCTypeDoc . '.spreadsheetml.worksheet+xml"/>' . XML_EOL;
        }

        if ($shouldUseInlineStrings) {
            $contentTypesXmlFileContents .= <<<EOD
  <Override PartName="/xl/sharedStrings.xml" ContentType="{$this->ooCTypeDoc}.spreadsheetml.sharedStrings+xml"/>

EOD;
        }
        $contentTypesXmlFileContents .= <<<EOD
  <Override PartName="/xl/styles.xml" ContentType="{$this->ooCTypeDoc}.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="{$this->ooCTypePkg}.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="{$this->ooCTypeDoc}.extended-properties+xml"/>
</Types>
EOD;

        $this->createFileWithContents($this->rootFolder, self::CONTENT_TYPES_XML_FILE_NAME, $contentTypesXmlFileContents);
        return $this;
    }

    /**
     * Creates the "workbook.xml" file under the "xl" folder
     *
     * @param Worksheet[] $worksheets
     * @return FileSystemHelper
     */
    public function createWorkbookFile($worksheets)
    {
        $workbookXmlFileContents = <<<EOD
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="{$this->ooXmlSheetDocNs}/main" xmlns:r="{$this->ooXmlOfficeDocNs}/relationships">
  <sheets>

EOD;

        /** @noinspection PhpUnnecessaryFullyQualifiedNameInspection */
        $escaper = \Box\Spout\Common\Escaper\XLSX::getInstance();

        /** @var Worksheet $worksheet */
        foreach ($worksheets as $worksheet) {
            $worksheetName = $worksheet->getExternalSheet()->getName();
            $worksheetId = $worksheet->getId();
            $rIdSheet = $worksheet->getSheetRId();
            $sheetName = $escaper->escape($worksheetName);
            $workbookXmlFileContents .= "    <sheet name=\"$sheetName\" sheetId=\"$worksheetId\" r:id=\"$rIdSheet\"/>" . XML_EOL;
        }

        $workbookXmlFileContents .= <<<EOD
  </sheets>
</workbook>
EOD;

        $this->createFileWithContents($this->xlFolder, self::WORKBOOK_XML_FILE_NAME, $workbookXmlFileContents);
        return $this;
    }

    /**
     * Creates the "workbook.xml.res" file under the "xl/_res" folder
     *
     * @param Worksheet[] $worksheets
     * @return FileSystemHelper
     */
    public function createWorkbookRelsFile($worksheets)
    {
        // NB: The "Target" filename is relative to "xl"; so Target=xl/styles.xml would be wrong.
        $workbookRelsXmlFileContents = <<<EOD
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="{$this->ooXmlPackageNs}/relationships">
  <Relationship Id="rIdStyles" Target="styles.xml" Type="{$this->ooXmlOfficeDocNs}/relationships/styles"/>

EOD;
        if ($shouldUseInlineStrings) {
            $workbookRelsXmlFileContents .=
              "  <Relationship Id=\"rIdSharedStrings\" Target=\"sharedStrings.xml\" Type=\"{$this->ooXmlOfficeDocNs}/relationships/sharedStrings\"/>" . XML_EOL;
        }
        /** @var Worksheet $worksheet */
        foreach ($worksheets as $worksheet) {
            $worksheetId = $worksheet->getId();
            $rId = $worksheet->getSheetRId();
            $workbookRelsXmlFileContents .=
              "  <Relationship Id=\"$rId\" Target=\"worksheets/sheet$worksheetId.xml\" Type=\"{$this->ooXmlOfficeDocNs}/relationships/worksheet\"/>" . XML_EOL;
        }

        $workbookRelsXmlFileContents .= '</Relationships>' . XML_EOL;

        $this->createFileWithContents($this->xlRelsFolder, self::WORKBOOK_RELS_XML_FILE_NAME, $workbookRelsXmlFileContents);
        return $this;
    }

    /**
     * Creates the "styles.xml" file under the "xl" folder
     *
     * @param StyleHelper $styleHelper
     * @return FileSystemHelper
     */
    public function createStylesFile($styleHelper)
    {
        $stylesXmlFileContents = $styleHelper->getStylesXMLFileContent();
        $this->createFileWithContents($this->xlFolder, self::STYLES_XML_FILE_NAME, $stylesXmlFileContents);
        return $this;
    }
}
