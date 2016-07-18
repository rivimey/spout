<?php

namespace Box\Spout\Autoloader;

require_once 'Psr4Autoloader.php';

/**
 * @var string $srcBaseDirectory
 * Full path to "src/Spout" which is what we want "Box\Spout" to map to.
 */
$srcBaseDirectory = __DIR__ . '/src/Spout/';

$loader = new Psr4Autoloader();
$loader->register();
$loader->addNamespace('Box\Spout', $srcBaseDirectory);
