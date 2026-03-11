-- MySQL dump 10.13  Distrib 8.0.19, for Win64 (x86_64)
--
-- Host: localhost    Database: natan291_korsys
-- ------------------------------------------------------
-- Server version	8.0.44

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!50503 SET NAMES utf8mb4 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `setores`
--

DROP TABLE IF EXISTS `setores`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `setores` (
  `set_id` int NOT NULL AUTO_INCREMENT,
  `set_nome` varchar(100) CHARACTER SET utf8mb3 COLLATE utf8mb3_unicode_ci NOT NULL,
  `set_descricao` varchar(200) CHARACTER SET utf8mb3 COLLATE utf8mb3_unicode_ci NOT NULL,
  `set_status` char(1) CHARACTER SET utf8mb3 COLLATE utf8mb3_unicode_ci NOT NULL DEFAULT 'A',
  `set_data_cadastro` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (`set_id`)
) ENGINE=InnoDB AUTO_INCREMENT=12 DEFAULT CHARSET=utf8mb3 COLLATE=utf8mb3_unicode_ci;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `setores`
--

LOCK TABLES `setores` WRITE;
/*!40000 ALTER TABLE `setores` DISABLE KEYS */;
INSERT INTO `setores` VALUES (1,'NÃO DEFINIDO','NÃO DEFINIDO','0','2026-02-19 18:47:22'),(2,'Tecnologia','Setor tecnologico','0','2026-02-19 18:52:27'),(7,'beleza',' ta funcionando','1','2026-03-03 04:16:27'),(8,'testando dnv','teste','0','2026-03-03 04:18:37'),(9,'agora ta funcionand','o mesmo','1','2026-03-03 04:18:59'),(10,' ultimo teste','mentira talvez não seja o ultimo','0','2026-03-03 04:19:53'),(11,'é o ultimo sim','quer dizer, não mais','1','2026-03-03 04:20:23');
/*!40000 ALTER TABLE `setores` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `usuarios`
--

DROP TABLE IF EXISTS `usuarios`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `usuarios` (
  `id` int NOT NULL AUTO_INCREMENT,
  `nome` varchar(120) CHARACTER SET utf8mb3 COLLATE utf8mb3_unicode_ci NOT NULL,
  `email` varchar(150) CHARACTER SET utf8mb3 COLLATE utf8mb3_unicode_ci NOT NULL,
  `senha` varchar(255) CHARACTER SET utf8mb3 COLLATE utf8mb3_unicode_ci DEFAULT NULL,
  `set_id` int NOT NULL DEFAULT '1',
  `nivel` enum('ADMIN','USER') CHARACTER SET utf8mb3 COLLATE utf8mb3_unicode_ci NOT NULL DEFAULT 'USER',
  `ativo` tinyint(1) NOT NULL DEFAULT '1',
  `criadoEm` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=141 DEFAULT CHARSET=utf8mb3 COLLATE=utf8mb3_unicode_ci;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `usuarios`
--

LOCK TABLES `usuarios` WRITE;
/*!40000 ALTER TABLE `usuarios` DISABLE KEYS */;
INSERT INTO `usuarios` VALUES (21,'natanael','natanael@gmail.com','123456',2,'ADMIN',1,'2026-02-03 19:32:31'),(22,'admin','admin@sistema.com','admin123',3,'ADMIN',1,'2026-02-03 19:32:31'),(23,'suporte','suporte@sistema.com','suporte123',1,'ADMIN',1,'2026-02-03 19:32:31'),(24,'joao.silva','joao.silva@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(25,'maria.oliveira','maria.oliveira@email.com','123456',1,'USER',0,'2026-02-03 19:32:31'),(26,'carlos.santos','carlos.santos@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(27,'ana.pereira','ana.pereira@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(28,'fernando.lima','fernando.lima@email.com','123456',1,'USER',0,'2026-02-03 19:32:31'),(29,'juliana.costa','juliana.costa@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(30,'roberto.alves','roberto.alves@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(31,'usuario01','usuario01@email.com',NULL,1,'USER',1,'2026-02-03 19:32:31'),(32,'usuario02','usuario02@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(33,'usuario03','usuario03@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(34,'usuario04','usuario04@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(35,'usuario05','usuario05@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(36,'usuario06','usuario06@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(37,'usuario07','usuario07@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(38,'usuario08','usuario08@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(39,'usuario09','usuario09@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(40,'usuario10','usuario10@email.com','123456',1,'USER',0,'2026-02-03 19:32:31'),(41,'usuario11','usuario11@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(42,'usuario42','email@teste.com.br','1123',1,'USER',1,'2026-02-03 19:32:31'),(43,'usuario13','usuario13@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(44,'usuario14','usuario14@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(45,'usuario15','usuario15@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(46,'usuario16','usuario16@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(47,'usuario17','usuario17@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(48,'usuario18','usuario18@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(49,'usuario19','usuario19@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(50,'usuario20','usuario20@email.com','123456',1,'USER',0,'2026-02-03 19:32:31'),(51,'usuario21','usuario21@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(52,'usuario22','usuario22@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(53,'usuario23','usuario23@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(54,'usuario24','usuario24@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(55,'usuario25','usuario25@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(56,'usuario26','usuario26@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(57,'usuario27','usuario27@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(58,'usuario28','usuario28@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(59,'usuario29','usuario29@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(60,'usuario30','usuario30@email.com','123456',1,'USER',0,'2026-02-03 19:32:31'),(61,'usuario31','usuario31@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(62,'usuario32','usuario32@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(63,'usuario33','usuario33@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(64,'usuario34','usuario34@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(65,'usuario35','usuario35@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(66,'usuario36','usuario36@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(67,'usuario37','usuario37@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(68,'usuario38','usuario38@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(69,'usuario39','usuario39@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(70,'usuario40','usuario40@email.com','123456',1,'USER',0,'2026-02-03 19:32:31'),(71,'usuario41','usuario41@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(72,'usuario42','usuario42@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(73,'usuario43','usuario43@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(74,'usuario74','usuario44@email.com','123456',1,'ADMIN',0,'2026-02-03 19:32:31'),(75,'usuario45','usuario45@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(76,'usuario46','usuario46@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(77,'usuario47','usuario47@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(78,'usuario48','usuario48@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(79,'usuario49','usuario49@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(80,'usuario50','usuario50@email.com','123456',1,'USER',0,'2026-02-03 19:32:31'),(81,'usuario51','usuario51@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(82,'usuario52','usuario52@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(83,'usuario53','usuario53@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(84,'usuario54','usuario54@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(85,'usuario55','usuario55@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(86,'usuario56','usuario56@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(87,'usuario57','usuario57@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(88,'usuario58','usuario58@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(89,'usuario59','usuario59@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(90,'usuario60','usuario60@email.com','123456',1,'USER',0,'2026-02-03 19:32:31'),(91,'usuario61','usuario61@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(92,'usuario62','usuario62@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(93,'usuario63','usuario63@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(94,'usuario64','usuario64@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(95,'usuario65','usuario65@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(96,'usuario66','usuario66@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(97,'usuario67','usuario67@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(98,'usuario68','usuario68@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(99,'usuario69','usuario69@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(100,'usuario70','usuario70@email.com','123456',1,'USER',0,'2026-02-03 19:32:31'),(101,'usuario71','usuario71@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(102,'usuario72','usuario72@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(103,'usuario73','usuario73@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(104,'usuario74','usuario74@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(105,'usuario75','usuario75@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(106,'usuario76','usuario76@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(107,'usuario77','usuario77@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(108,'usuario78','usuario78@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(109,'usuario79','usuario79@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(110,'usuario80','usuario80@email.com','123456',1,'USER',0,'2026-02-03 19:32:31'),(111,'usuario81','usuario81@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(112,'usuario82','usuario82@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(113,'usuario83','usuario83@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(114,'usuario84','usuario84@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(115,'usuario85','usuario85@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(116,'usuario86','usuario86@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(117,'usuario87','usuario87@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(118,'usuario88','usuario88@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(119,'usuario89','usuario89@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(120,'usuario90','usuario90@email.com','123456',1,'USER',0,'2026-02-03 19:32:31'),(121,'usuario91','usuario91@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(122,'usuario92','usuario92@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(123,'usuario93','usuario93@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(124,'usuario94','usuario94@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(125,'usuario95','usuario95@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(126,'usuario96','usuario96@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(127,'usuario97','usuario97@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(128,'usuario98','usuario98@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(129,'usuario99','usuario99@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(130,'usuario100','usuario100@email.com','123456',1,'USER',1,'2026-02-03 19:32:31'),(132,'Nathan','nathanjorge@gmail.com','12345678',1,'ADMIN',1,'2026-02-03 19:52:28'),(134,'jorge','jorge1@gmail.com','senhajorge',1,'USER',1,'2026-02-10 19:41:50'),(135,'jorginho','jorginemail@gmail','jorgin123',1,'USER',0,'2026-02-10 19:53:41'),(136,'fidojorge','fidojorge@gmail','fidojorgin',1,'ADMIN',1,'2026-02-10 19:56:21'),(137,'robson','robson@gmail.com','robinho',1,'USER',0,'2026-02-10 20:02:51'),(138,'jorel','joreek@jorela.com.br','teste',1,'USER',1,'2026-02-18 18:13:24'),(139,'pão','paozin@gmail.com','111444',1,'ADMIN',1,'2026-02-18 18:42:13');
/*!40000 ALTER TABLE `usuarios` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `usuarios_perms`
--

DROP TABLE IF EXISTS `usuarios_perms`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `usuarios_perms` (
  `id` int NOT NULL AUTO_INCREMENT,
  `id_user` int NOT NULL,
  `modulo` varchar(77) CHARACTER SET utf8mb3 COLLATE utf8mb3_unicode_ci NOT NULL,
  `permitido` tinyint(1) NOT NULL DEFAULT '0',
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=9 DEFAULT CHARSET=utf8mb3 COLLATE=utf8mb3_unicode_ci;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `usuarios_perms`
--

LOCK TABLES `usuarios_perms` WRITE;
/*!40000 ALTER TABLE `usuarios_perms` DISABLE KEYS */;
INSERT INTO `usuarios_perms` VALUES (1,1,'DEUSJORGE',1),(2,21,'CADASTRO DE CLIENTES',1),(3,22,'CADASTRO DE CLIENTES',0),(4,23,'CADASTRO DE CLIENTES',1),(5,34,'CADASTRO DE CLIENTES',1),(6,30,'CADASTRO DE CLIENTES',1),(7,132,'CADASTRO DE CLIENTES',1),(8,132,'CADASTRO DE USUARIOS',1);
/*!40000 ALTER TABLE `usuarios_perms` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Dumping routines for database 'natan291_korsys'
--
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2026-03-11  1:58:33
