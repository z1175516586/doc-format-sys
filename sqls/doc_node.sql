/*
 Navicat Premium Dump SQL

 Source Server         : 172.23.92.18
 Source Server Type    : MySQL
 Source Server Version : 80025 (8.0.25)
 Source Host           : localhost:3306
 Source Schema         : szzs

 Target Server Type    : MySQL
 Target Server Version : 80025 (8.0.25)
 File Encoding         : 65001

 Date: 27/10/2026 21:07:07
*/

SET NAMES utf8mb4;
SET FOREIGN_KEY_CHECKS = 0;

-- ----------------------------
-- Table structure for doc_node
-- ----------------------------
DROP TABLE IF EXISTS `doc_node`;
CREATE TABLE `doc_node`  (
  `id` bigint NOT NULL AUTO_INCREMENT COMMENT '节点ID',
  `node_code` varchar(22) CHARACTER SET utf8mb4 COLLATE utf8mb4_0900_ai_ci NOT NULL COMMENT '当前文件唯一ID',
  `root_id` bigint NOT NULL COMMENT '根节点ID(文档族ID)',
  `parent_id` bigint NULL DEFAULT NULL COMMENT '父节点ID',
  `file_path` varchar(500) CHARACTER SET utf8mb4 COLLATE utf8mb4_0900_ai_ci NULL DEFAULT NULL COMMENT '文件路径',
  `created_at` datetime NULL DEFAULT CURRENT_TIMESTAMP COMMENT '创建时间',
  PRIMARY KEY (`id`) USING BTREE,
  UNIQUE INDEX `uk_node_code`(`node_code` ASC) USING BTREE,
  INDEX `idx_root_id`(`root_id` ASC) USING BTREE,
  INDEX `idx_parent_id`(`parent_id` ASC) USING BTREE,
  INDEX `idx_created_at`(`created_at` ASC) USING BTREE
) ENGINE = InnoDB AUTO_INCREMENT = 261 CHARACTER SET = utf8mb4 COLLATE = utf8mb4_0900_ai_ci COMMENT = '文档版本节点表' ROW_FORMAT = DYNAMIC;

SET FOREIGN_KEY_CHECKS = 1;
