<?php
function plugin_version_driverslist()
{
return array('name' => 'driverslist',
'version' => '1.1',
'author'=> 'Stephane PAUTREL',
'license' => 'GPLv2',
'verMinOcs' => '2.2');
}

function plugin_init_driverslist()
{
$object = new plugins;
$object -> add_cd_entry("driverslist","hardware");

$object -> sql_query("CREATE TABLE IF NOT EXISTS `driverslist` (
  `ID` INT(11) NOT NULL AUTO_INCREMENT,
  `HARDWARE_ID` INT(11) NOT NULL,
  `CLASS` VARCHAR(255) DEFAULT NULL,
  `DESCRIPTION` VARCHAR(255) DEFAULT NULL,
  `NAME` VARCHAR(255) DEFAULT NULL,
  `PROVIDERNAME` VARCHAR(255) DEFAULT NULL,
  `VERSION` VARCHAR(255) DEFAULT NULL,
  `DATE` VARCHAR(255) DEFAULT NULL,
  `INFNAME` VARCHAR(255) DEFAULT NULL,
  `FRIENDLYNAME` VARCHAR(255) DEFAULT NULL,
  `SIGNED` VARCHAR(255) DEFAULT NULL,
  `DEVICEID` VARCHAR(255) DEFAULT NULL,
  PRIMARY KEY  (`ID`,`HARDWARE_ID`)
) ENGINE=INNODB ;");

}

function plugin_delete_driverslist()
{
$object = new plugins;
$object -> del_cd_entry("driverslist");
$object -> sql_query("DROP TABLE `driverslist`;");

}

?>
