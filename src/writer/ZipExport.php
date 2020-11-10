<?php


namespace data\writer;
use RecursiveDirectoryIterator;
use RecursiveIteratorIterator;
use ZipArchive;

class ZipExport
{
    /**
     * zip打包
     * @param $folder
     * @param $zipFile
     */
    public static function zip($folder, $zipFile)
    {
        // Get real path for our folder
        $rootPath = realpath($folder);

        // Initialize archive object
        $zip = new ZipArchive();
        $zip->open($zipFile, ZipArchive::CREATE | ZipArchive::OVERWRITE);

        // Create recursive directory iterator
        $files = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($rootPath), RecursiveIteratorIterator::LEAVES_ONLY);

        foreach ($files as $name => $file) {
            // Skip directories (they would be added automatically)
            if (!$file->isDir()) {
                // Get real and relative path for current file
                $filePath = $file->getRealPath();
                $relativePath = substr($filePath, strlen($rootPath) + 1);

                // Add current file to archive
                $zip->addFile($filePath, $relativePath);
            }
        }

        // Zip archive will be created only after closing object
        $zip->close();
    }
}