-- --------------------------------------------------------
-- 호스트:                          127.0.0.1
-- 서버 버전:                        10.4.18-MariaDB - mariadb.org binary distribution
-- 서버 OS:                        Win64
-- HeidiSQL 버전:                  11.0.0.5919
-- --------------------------------------------------------

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET NAMES utf8 */;
/*!50503 SET NAMES utf8mb4 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;


-- op_system 데이터베이스 구조 내보내기
DROP DATABASE IF EXISTS `op_system`;
CREATE DATABASE IF NOT EXISTS `op_system` /*!40100 DEFAULT CHARACTER SET utf8 */;
USE `op_system`;

-- 테이블 op_system.a_authority 구조 내보내기
DROP TABLE IF EXISTS `a_authority`;
CREATE TABLE IF NOT EXISTS `a_authority` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `authority` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=12 DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.a_auth_table 구조 내보내기
DROP TABLE IF EXISTS `a_auth_table`;
CREATE TABLE IF NOT EXISTS `a_auth_table` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `user_id` int(11) NOT NULL,
  `authority_id` int(11) NOT NULL,
  PRIMARY KEY (`id`) USING BTREE,
  KEY `user_id` (`user_id`),
  KEY `ovs_dept` (`authority_id`) USING BTREE,
  CONSTRAINT `FK_a_auth_table_a_authority` FOREIGN KEY (`authority_id`) REFERENCES `a_authority` (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=185 DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.a_branch_admin 구조 내보내기
DROP TABLE IF EXISTS `a_branch_admin`;
CREATE TABLE IF NOT EXISTS `a_branch_admin` (
  `church_sid` varchar(30) NOT NULL COMMENT '교회코드',
  `Branch_cd` varchar(50) NOT NULL COMMENT '지교회번호',
  `church_gb` varchar(5) NOT NULL COMMENT '교회구분',
  `church_nm_ko` varchar(100) NOT NULL COMMENT '한글 지교회명',
  `church_nm_en` varchar(100) NOT NULL COMMENT '영문 지교회명',
  `church_nm_lo` varchar(100) DEFAULT NULL COMMENT '현지 지교회명',
  `suspend` varchar(5) NOT NULL DEFAULT '0' COMMENT '운영여부',
  `check` varchar(5) DEFAULT NULL COMMENT '확인',
  `main_church` varchar(100) NOT NULL COMMENT '관리 교회',
  `church_establish` varchar(100) NOT NULL COMMENT '분가시킨교회',
  `main_branch` varchar(100) DEFAULT NULL COMMENT '관리 지교회',
  `start_dt` date NOT NULL COMMENT '시작일',
  `end_dt` date DEFAULT NULL COMMENT '종료일',
  `input_dt` varchar(50) NOT NULL DEFAULT '0' COMMENT '입력일',
  `start_atten` varchar(50) NOT NULL DEFAULT '0' COMMENT '최초 시작성도',
  `Once_all` varchar(50) DEFAULT NULL COMMENT '전월 1회출석',
  `Once_stu` varchar(50) DEFAULT NULL COMMENT '전월 1회출석(학)',
  `wmc_media` varchar(50) DEFAULT NULL,
  `continent` varchar(20) NOT NULL DEFAULT '0' COMMENT '대륙',
  `country` varchar(30) NOT NULL DEFAULT '0' COMMENT '국가',
  `admin1` varchar(50) NOT NULL DEFAULT '0',
  `admin2` varchar(50) DEFAULT '0',
  `admin3` varchar(50) DEFAULT '0',
  `admin4` varchar(50) DEFAULT '0',
  `geo_cd` int(10) NOT NULL DEFAULT 0 COMMENT 'GEO코드',
  `address` varchar(300) DEFAULT '0' COMMENT '주소',
  `manager` varchar(100) DEFAULT '0' COMMENT '관리자',
  `lifeno` varchar(30) DEFAULT '0' COMMENT '관리자 생명번호',
  `title` varchar(10) DEFAULT '0' COMMENT '직분',
  `position` varchar(15) DEFAULT '0' COMMENT '직책',
  `age_group` varchar(15) DEFAULT '0',
  `appoint_dt` varchar(10) DEFAULT NULL COMMENT '임명일',
  `baptism` varchar(5) DEFAULT '0' COMMENT '침례권',
  PRIMARY KEY (`church_sid`),
  KEY `country` (`country`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.a_churchlist_admin 구조 내보내기
DROP TABLE IF EXISTS `a_churchlist_admin`;
CREATE TABLE IF NOT EXISTS `a_churchlist_admin` (
  `church_sid` varchar(15) NOT NULL COMMENT '교회코드',
  `division` varchar(15) NOT NULL,
  `union` varchar(15) NOT NULL COMMENT '대회',
  `association` varchar(15) NOT NULL COMMENT '연합회',
  `church_nm_ko` varchar(50) NOT NULL COMMENT '한글 교회명',
  `church_nm_en` varchar(80) NOT NULL COMMENT '영문 교회명',
  `view_photo_church` varchar(10) DEFAULT NULL COMMENT '사진보기',
  `manager_nm` varchar(50) DEFAULT NULL COMMENT '관리자명',
  `view_photo_maganer` varchar(10) DEFAULT NULL COMMENT '관리자 사진보기',
  `title` varchar(15) DEFAULT NULL COMMENT '직분',
  `position` varchar(30) DEFAULT NULL COMMENT '직책',
  `church_gb` varchar(10) NOT NULL COMMENT '교회형태',
  `baptism` varchar(5) DEFAULT NULL COMMENT '침례권',
  `main_church` varchar(50) DEFAULT NULL COMMENT '관리교회',
  `main_branch` varchar(50) DEFAULT NULL COMMENT '관리 지교회',
  `church_num` varchar(20) NOT NULL COMMENT '교회번호',
  `church_cd` int(8) NOT NULL DEFAULT 0 COMMENT '교회코드',
  `atten_visit` varchar(5) NOT NULL COMMENT '방문 출석 가능여부',
  `geo_cd` int(10) NOT NULL DEFAULT 0 COMMENT 'GEO코드',
  `continent` varchar(15) NOT NULL COMMENT '대륙',
  `country` varchar(30) NOT NULL COMMENT '국가',
  `admin1` varchar(30) NOT NULL,
  `city` varchar(50) DEFAULT NULL COMMENT '도시',
  `city_population` varchar(10) DEFAULT NULL COMMENT '도시인구',
  `latitude` decimal(15,10) NOT NULL DEFAULT 0.0000000000 COMMENT '위도',
  `longitude` decimal(15,10) NOT NULL DEFAULT 0.0000000000 COMMENT '경도',
  PRIMARY KEY (`church_sid`),
  KEY `country` (`country`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.a_counsel_category 구조 내보내기
DROP TABLE IF EXISTS `a_counsel_category`;
CREATE TABLE IF NOT EXISTS `a_counsel_category` (
  `counsel_category` varchar(50) NOT NULL DEFAULT ''
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.a_dimdate 구조 내보내기
DROP TABLE IF EXISTS `a_dimdate`;
CREATE TABLE IF NOT EXISTS `a_dimdate` (
  `DateKey` int(11) NOT NULL,
  `DateValue` date NOT NULL,
  `NextDayValue` date NOT NULL,
  `YearValue` smallint(6) NOT NULL,
  `YearQuarter` int(11) NOT NULL,
  `YearMonth` int(11) NOT NULL,
  `YearDayOfYear` int(11) NOT NULL,
  `QuarterValue` tinyint(4) NOT NULL,
  `MonthValue` tinyint(4) NOT NULL,
  `DayOfYear` smallint(6) NOT NULL,
  `DayOfMonth` smallint(6) NOT NULL,
  `DayOfWeek` tinyint(4) NOT NULL,
  `MonthName` varchar(3) NOT NULL,
  `MonthNameLong` varchar(9) NOT NULL,
  `WeekdayName` varchar(3) NOT NULL,
  `WeekDayNameLong` varchar(9) NOT NULL,
  `StartOfYearDate` date NOT NULL,
  `EndOfYearDate` date NOT NULL,
  `StartOfQuarterDate` date NOT NULL,
  `EndOfQuarterDate` date NOT NULL,
  `StartOfMonthDate` date NOT NULL,
  `EndOfMonthDate` date NOT NULL,
  `StartOfWeekStartingSunDate` date NOT NULL,
  `EndOfWeekStartingSunDate` date NOT NULL,
  PRIMARY KEY (`DateValue`),
  KEY `DateKey` (`DateKey`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.a_position 구조 내보내기
DROP TABLE IF EXISTS `a_position`;
CREATE TABLE IF NOT EXISTS `a_position` (
  `position` varchar(50) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.a_position2 구조 내보내기
DROP TABLE IF EXISTS `a_position2`;
CREATE TABLE IF NOT EXISTS `a_position2` (
  `position2` varchar(50) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.a_position_spouse 구조 내보내기
DROP TABLE IF EXISTS `a_position_spouse`;
CREATE TABLE IF NOT EXISTS `a_position_spouse` (
  `position` varchar(10) NOT NULL COMMENT '남편직책',
  `position_Spouse` varchar(10) NOT NULL COMMENT '사모직책'
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.a_title 구조 내보내기
DROP TABLE IF EXISTS `a_title`;
CREATE TABLE IF NOT EXISTS `a_title` (
  `title` varchar(50) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.a_union 구조 내보내기
DROP TABLE IF EXISTS `a_union`;
CREATE TABLE IF NOT EXISTS `a_union` (
  `union_cd` int(7) NOT NULL AUTO_INCREMENT COMMENT '연합회코드',
  `union_nm` varchar(10) NOT NULL DEFAULT '0' COMMENT '연합회명',
  `suspend` int(11) NOT NULL DEFAULT 0 COMMENT '논리삭제',
  `ovs_dept` int(3) NOT NULL DEFAULT 0 COMMENT '관리부서',
  `sort_order` int(3) NOT NULL DEFAULT 0 COMMENT '정렬순서',
  PRIMARY KEY (`union_cd`)
) ENGINE=InnoDB AUTO_INCREMENT=135 DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.a_visa 구조 내보내기
DROP TABLE IF EXISTS `a_visa`;
CREATE TABLE IF NOT EXISTS `a_visa` (
  `visa_id` int(11) NOT NULL AUTO_INCREMENT,
  `visa_nm` char(50) NOT NULL DEFAULT '0',
  PRIMARY KEY (`visa_id`)
) ENGINE=InnoDB AUTO_INCREMENT=15 DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_attendance 구조 내보내기
DROP TABLE IF EXISTS `db_attendance`;
CREATE TABLE IF NOT EXISTS `db_attendance` (
  `church_sid` varchar(50) NOT NULL COMMENT '교회코드',
  `attendance_dt` date NOT NULL COMMENT '출석월',
  `once_all` int(6) NOT NULL DEFAULT 0 COMMENT '전체1회',
  `forth_all` int(6) NOT NULL DEFAULT 0 COMMENT '전체4회',
  `once_stu` int(6) NOT NULL DEFAULT 0 COMMENT '학생이상1회',
  `forth_stu` int(6) NOT NULL DEFAULT 0 COMMENT '학생이상4회',
  `tithe_all` int(6) NOT NULL DEFAULT 0 COMMENT '전체 반차',
  `tithe_stu` int(6) NOT NULL DEFAULT 0 COMMENT '학생이상 반차',
  `baptism_all` int(6) NOT NULL DEFAULT 0 COMMENT '전체 침례',
  `evangelist` int(6) NOT NULL DEFAULT 0 COMMENT '고정전도인',
  `gl` int(6) NOT NULL DEFAULT 0 COMMENT '지역장',
  `ul` int(6) NOT NULL DEFAULT 0 COMMENT '구역장',
  PRIMARY KEY (`church_sid`,`attendance_dt`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_attendance_backup 구조 내보내기
DROP TABLE IF EXISTS `db_attendance_backup`;
CREATE TABLE IF NOT EXISTS `db_attendance_backup` (
  `church_sid` varchar(50) NOT NULL COMMENT '교회코드',
  `attendance_dt` date NOT NULL COMMENT '출석월',
  `once_all` int(6) NOT NULL DEFAULT 0 COMMENT '전체1회',
  `forth_all` int(6) NOT NULL DEFAULT 0 COMMENT '전체4회',
  `once_stu` int(6) NOT NULL DEFAULT 0 COMMENT '학생이상1회',
  `forth_stu` int(6) NOT NULL DEFAULT 0 COMMENT '학생이상4회',
  `tithe_all` int(6) NOT NULL DEFAULT 0 COMMENT '전체 반차',
  `tithe_stu` int(6) NOT NULL DEFAULT 0 COMMENT '학생이상 반차',
  `baptism_all` int(6) NOT NULL DEFAULT 0 COMMENT '전체 침례',
  `evangelist` int(6) NOT NULL DEFAULT 0 COMMENT '고정전도인',
  `gl` int(6) NOT NULL DEFAULT 0 COMMENT '지역장',
  `ul` int(6) NOT NULL DEFAULT 0 COMMENT '구역장',
  PRIMARY KEY (`church_sid`,`attendance_dt`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_branchleader 구조 내보내기
DROP TABLE IF EXISTS `db_branchleader`;
CREATE TABLE IF NOT EXISTS `db_branchleader` (
  `bcleader_cd` int(11) NOT NULL AUTO_INCREMENT,
  `church_sid` varchar(15) NOT NULL COMMENT '교회코드',
  `Start_dt` date NOT NULL COMMENT '시작일',
  `End_dt` date NOT NULL DEFAULT '9999-12-31' COMMENT '종료일',
  `lifeno` varchar(50) NOT NULL COMMENT '생명번호',
  `responsibility` varchar(30) NOT NULL DEFAULT '관리자' COMMENT '관리자/단순소속',
  PRIMARY KEY (`bcleader_cd`),
  KEY `church_sid` (`church_sid`),
  KEY `Start_dt` (`Start_dt`,`End_dt`),
  KEY `lifeno` (`lifeno`)
) ENGINE=InnoDB AUTO_INCREMENT=11028 DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_churchlist 구조 내보내기
DROP TABLE IF EXISTS `db_churchlist`;
CREATE TABLE IF NOT EXISTS `db_churchlist` (
  `church_sid` varchar(20) NOT NULL COMMENT '교회코드',
  `church_nm` varchar(50) NOT NULL COMMENT '교회명',
  `church_gb` varchar(15) NOT NULL COMMENT '교회형태',
  `manager_cd` varchar(100) DEFAULT NULL COMMENT '관리자 생명번호',
  `main_church_cd` varchar(20) DEFAULT NULL COMMENT '관리교회 코드',
  `start_dt` date DEFAULT NULL COMMENT '시작일',
  `end_dt` date DEFAULT NULL COMMENT '종료일',
  `ovs_dept` varchar(15) DEFAULT NULL COMMENT '해외국 관리부서',
  `suspend` tinyint(1) NOT NULL DEFAULT 0 COMMENT '논리삭제',
  `sort_order` int(10) NOT NULL COMMENT '정렬순서',
  `geo_cd` int(10) DEFAULT NULL,
  PRIMARY KEY (`church_sid`),
  KEY `main_church_cd` (`main_church_cd`),
  KEY `geo_cd` (`geo_cd`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_churchlist_custom 구조 내보내기
DROP TABLE IF EXISTS `db_churchlist_custom`;
CREATE TABLE IF NOT EXISTS `db_churchlist_custom` (
  `church_sid` varchar(20) NOT NULL COMMENT '교회코드',
  `church_nm` varchar(50) NOT NULL COMMENT '교회이름',
  `church_gb` varchar(15) NOT NULL COMMENT '교회형태',
  `manager_cd` varchar(100) DEFAULT NULL COMMENT '관리자 생명번호',
  `main_church_cd` varchar(20) DEFAULT NULL COMMENT '관리교회 코드',
  `start_dt` date DEFAULT NULL COMMENT '시작일',
  `end_dt` date DEFAULT NULL COMMENT '종료일',
  `ovs_dept` varchar(15) DEFAULT NULL COMMENT '해외국 관리부서',
  `suspend` tinyint(1) NOT NULL DEFAULT 0 COMMENT '논리삭제',
  `sort_order` int(10) NOT NULL COMMENT '정렬순서',
  `geo_cd` int(10) DEFAULT NULL,
  PRIMARY KEY (`church_sid`),
  KEY `main_church_cd` (`main_church_cd`),
  KEY `geo_cd` (`geo_cd`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_church_map 구조 내보내기
DROP TABLE IF EXISTS `db_church_map`;
CREATE TABLE IF NOT EXISTS `db_church_map` (
  `sid` varchar(50) NOT NULL,
  `map` longtext DEFAULT NULL,
  PRIMARY KEY (`sid`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_counsel 구조 내보내기
DROP TABLE IF EXISTS `db_counsel`;
CREATE TABLE IF NOT EXISTS `db_counsel` (
  `counsel_id` int(11) NOT NULL AUTO_INCREMENT,
  `life_no` varchar(50) NOT NULL,
  `counsel_dt` date NOT NULL DEFAULT curdate(),
  `category` varchar(50) NOT NULL DEFAULT '',
  `title` varchar(200) DEFAULT NULL,
  `content` text DEFAULT '',
  `result` text DEFAULT '',
  `remark` text DEFAULT '',
  `status` varchar(50) NOT NULL DEFAULT '',
  PRIMARY KEY (`counsel_id`)
) ENGINE=InnoDB AUTO_INCREMENT=25 DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_country 구조 내보내기
DROP TABLE IF EXISTS `db_country`;
CREATE TABLE IF NOT EXISTS `db_country` (
  `ctry_nm` varchar(70) NOT NULL COMMENT '국가명(한글)',
  `ctry_nm_en` varchar(70) NOT NULL COMMENT '국가명(영어)',
  `population` int(11) NOT NULL DEFAULT 0 COMMENT '인구수',
  PRIMARY KEY (`ctry_nm`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_division 구조 내보내기
DROP TABLE IF EXISTS `db_division`;
CREATE TABLE IF NOT EXISTS `db_division` (
  `geo_code` int(8) NOT NULL,
  `division` varchar(50) DEFAULT NULL,
  `association` varchar(50) DEFAULT NULL,
  `region` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`geo_code`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_familyinfo 구조 내보내기
DROP TABLE IF EXISTS `db_familyinfo`;
CREATE TABLE IF NOT EXISTS `db_familyinfo` (
  `family_id` int(5) NOT NULL AUTO_INCREMENT COMMENT '구성원id',
  `family_cd` int(5) NOT NULL COMMENT '가족코드',
  `relations` varchar(5) NOT NULL COMMENT '가족관계',
  `lifeno` varchar(20) DEFAULT NULL COMMENT '생명번호',
  `name_ko` varchar(50) DEFAULT NULL COMMENT '한글이름',
  `name_en` varchar(50) DEFAULT NULL COMMENT '영문이름',
  `church_sid` varchar(15) DEFAULT NULL COMMENT '교회',
  `title` varchar(10) DEFAULT NULL COMMENT '직분',
  `position` varchar(10) DEFAULT NULL COMMENT '직책',
  `birthday` date DEFAULT NULL COMMENT '생년월일',
  `education` varchar(20) DEFAULT NULL COMMENT '학력',
  `religion` varchar(20) DEFAULT NULL COMMENT '종교',
  `recognition` varchar(5) DEFAULT NULL COMMENT '본교인식',
  `memo` varchar(300) DEFAULT NULL COMMENT '메모',
  `suspend` tinyint(1) NOT NULL DEFAULT 0 COMMENT '0-생존,1-별세',
  PRIMARY KEY (`family_id`),
  KEY `family_cd` (`family_cd`),
  KEY `church_sid` (`church_sid`),
  KEY `lifeno` (`lifeno`)
) ENGINE=InnoDB AUTO_INCREMENT=17846 DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_flight_schedule 구조 내보내기
DROP TABLE IF EXISTS `db_flight_schedule`;
CREATE TABLE IF NOT EXISTS `db_flight_schedule` (
  `flight_cd` int(11) NOT NULL AUTO_INCREMENT,
  `lifeno` varchar(20) NOT NULL,
  `flight_dt` date NOT NULL,
  `departure` varchar(50) NOT NULL,
  `destination` varchar(50) NOT NULL,
  `visit_purpose` varchar(100) DEFAULT NULL,
  PRIMARY KEY (`flight_cd`),
  KEY `lifeno` (`lifeno`)
) ENGINE=InnoDB AUTO_INCREMENT=24330 DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_geodata 구조 내보내기
DROP TABLE IF EXISTS `db_geodata`;
CREATE TABLE IF NOT EXISTS `db_geodata` (
  `geo_cd` int(10) NOT NULL COMMENT 'geo_cd',
  `country_nm_ko` varchar(50) NOT NULL DEFAULT '' COMMENT '국가명(한글)',
  `country_population` int(20) NOT NULL DEFAULT 0 COMMENT '국가 인구수',
  `admin1_geo_cd` int(20) DEFAULT NULL COMMENT 'admin1 geo_cd',
  `admin1_nm_ko` varchar(100) DEFAULT NULL COMMENT 'admin1 명칭',
  `admin1_population` int(11) DEFAULT NULL COMMENT 'admin1 인구수',
  `admin2_geo_cd` int(20) DEFAULT NULL COMMENT 'admin2 geo_cd',
  `admin2_nm_ko` varchar(100) DEFAULT NULL COMMENT 'admin2 명칭',
  `admin2_population` int(11) DEFAULT NULL COMMENT 'admin2 인구수',
  `admin3_geo_cd` int(20) DEFAULT NULL COMMENT 'admin3 geo_cd',
  `admin3_nm_ko` varchar(100) DEFAULT NULL COMMENT 'admin3 명칭',
  `admin3_population` int(11) DEFAULT NULL COMMENT 'admin3 인구수',
  `admin4_geo_cd` int(20) DEFAULT NULL COMMENT 'admin4 geo_cd',
  `admin4_nm_ko` varchar(100) DEFAULT NULL COMMENT 'admin4 명칭',
  `admin4_population` int(11) DEFAULT NULL COMMENT 'admin4 인구수',
  `gospel_continent` int(10) DEFAULT NULL COMMENT '복음대륙',
  `gospel_country` int(10) DEFAULT NULL COMMENT '복음국가',
  `gospel_region` int(10) DEFAULT NULL COMMENT '복음지방',
  `gospel_city` int(10) DEFAULT NULL COMMENT '복음도시',
  `center` varchar(5) DEFAULT NULL COMMENT '센터도시',
  `latitude` decimal(15,10) DEFAULT NULL COMMENT '위도',
  `longitude` decimal(15,10) DEFAULT NULL COMMENT '경도',
  `mission_continent` varchar(50) DEFAULT NULL COMMENT '선교대룩_부서담당대륙',
  `mission_department` varchar(50) DEFAULT NULL COMMENT '선교대륙_복음대륙',
  `department` varchar(50) DEFAULT NULL COMMENT '담당부서',
  `division` varchar(50) DEFAULT NULL COMMENT '담당과',
  PRIMARY KEY (`geo_cd`),
  KEY `country_nm_ko` (`country_nm_ko`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COMMENT='Geo 전산 데이터';

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_history_church 구조 내보내기
DROP TABLE IF EXISTS `db_history_church`;
CREATE TABLE IF NOT EXISTS `db_history_church` (
  `his_cd` int(5) NOT NULL AUTO_INCREMENT,
  `church_sid` varchar(20) NOT NULL COMMENT '교회코드',
  `his_dt` date DEFAULT NULL COMMENT '이력일자',
  `history` varchar(200) DEFAULT NULL COMMENT '이력내용',
  PRIMARY KEY (`his_cd`),
  KEY `church_sid` (`church_sid`)
) ENGINE=InnoDB AUTO_INCREMENT=1259 DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_history_church_establish 구조 내보내기
DROP TABLE IF EXISTS `db_history_church_establish`;
CREATE TABLE IF NOT EXISTS `db_history_church_establish` (
  `church_esta_cd` int(6) NOT NULL AUTO_INCREMENT,
  `church_sid_custom` int(6) NOT NULL,
  `start_dt` date NOT NULL,
  `end_dt` date DEFAULT '9999-12-31',
  `church_sid` varchar(20) NOT NULL,
  PRIMARY KEY (`church_esta_cd`),
  KEY `start_dt` (`start_dt`,`end_dt`),
  KEY `church_sid_custom` (`church_sid_custom`),
  KEY `church_sid` (`church_sid`)
) ENGINE=InnoDB AUTO_INCREMENT=8964 DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_ovs_dept 구조 내보내기
DROP TABLE IF EXISTS `db_ovs_dept`;
CREATE TABLE IF NOT EXISTS `db_ovs_dept` (
  `dept_id` int(3) NOT NULL COMMENT '부서id',
  `dept_lv1` varchar(30) NOT NULL COMMENT '부서lv1',
  `dept_lv2` varchar(30) DEFAULT NULL COMMENT '부서lv2',
  `dept_lv3` varchar(30) DEFAULT NULL COMMENT '부서lv3',
  `dept_nm` varchar(30) NOT NULL COMMENT '부서명',
  `dept_phonecard` varchar(30) DEFAULT NULL COMMENT '국제전화카드',
  `dept_picpath` varchar(150) DEFAULT NULL COMMENT '사진경로',
  `sort_order` int(5) NOT NULL COMMENT '정렬기준',
  `suspended` tinyint(1) NOT NULL DEFAULT 0 COMMENT '논리삭제',
  PRIMARY KEY (`dept_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_passport_photo 구조 내보내기
DROP TABLE IF EXISTS `db_passport_photo`;
CREATE TABLE IF NOT EXISTS `db_passport_photo` (
  `lifeno` varchar(50) NOT NULL,
  `photo` longtext DEFAULT NULL,
  PRIMARY KEY (`lifeno`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_pastoralstaff 구조 내보내기
DROP TABLE IF EXISTS `db_pastoralstaff`;
CREATE TABLE IF NOT EXISTS `db_pastoralstaff` (
  `lifeno` varchar(20) NOT NULL COMMENT '생명번호',
  `name_ko` varchar(50) NOT NULL COMMENT '한글이름',
  `name_en` varchar(50) NOT NULL COMMENT '영문이름',
  `nationality` varchar(50) NOT NULL COMMENT '국적',
  `birthday` date NOT NULL COMMENT '생년월일',
  `phone` varchar(80) DEFAULT NULL COMMENT '전화번호',
  `lifeno_child1` varchar(20) DEFAULT NULL COMMENT '자녀1 생명번호',
  `name_ko_child1` varchar(50) DEFAULT NULL COMMENT '자녀1 한글이름',
  `name_en_child1` varchar(50) DEFAULT NULL COMMENT '자녀1 영문이름',
  `birthday_child1` date DEFAULT NULL COMMENT '자녀1 생년월일',
  `phone_child1` varchar(30) DEFAULT NULL COMMENT '자녀1 전화번호',
  `lifeno_child2` varchar(20) DEFAULT NULL COMMENT '자녀2 생명번호',
  `name_ko_child2` varchar(50) DEFAULT NULL COMMENT '자녀2 한글이름',
  `name_en_child2` varchar(50) DEFAULT NULL COMMENT '자녀2 영문이름',
  `birthday_child2` date DEFAULT NULL COMMENT '자녀2 생년월일',
  `phone_child2` varchar(30) DEFAULT NULL COMMENT '자녀2 전화번호',
  `lifeno_child3` varchar(20) DEFAULT NULL COMMENT '자녀2 생명번호',
  `name_ko_child3` varchar(50) DEFAULT NULL COMMENT '자녀2 한글이름',
  `name_en_child3` varchar(50) DEFAULT NULL COMMENT '자녀2 영문이름',
  `birthday_child3` date DEFAULT NULL COMMENT '자녀2 생년월일',
  `phone_child3` varchar(30) DEFAULT NULL COMMENT '자녀2 전화번호',
  `home` varchar(200) DEFAULT NULL COMMENT '본가',
  `family` varchar(700) DEFAULT NULL COMMENT '가족사항',
  `health` text DEFAULT NULL COMMENT '건강사항',
  `other` text DEFAULT NULL COMMENT '기타사항',
  `baptism` varchar(2) NOT NULL DEFAULT '무' COMMENT '침례권',
  `ordination_prayer` date DEFAULT NULL COMMENT '침례권안수',
  `appo_ovs` date DEFAULT NULL COMMENT '해외 최초 발령일',
  `wedding_dt` date DEFAULT NULL COMMENT '혼인일',
  `theological_order` int(4) DEFAULT NULL COMMENT '한국인 생도기수',
  `education` varchar(100) DEFAULT NULL COMMENT '학력',
  `salary` int(15) NOT NULL DEFAULT 0 COMMENT '유급',
  `suspend` tinyint(1) NOT NULL DEFAULT 0 COMMENT '논리삭제',
  `ovs_dept` int(3) NOT NULL DEFAULT 0 COMMENT '관리부서',
  PRIMARY KEY (`lifeno`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_pastoralstaff_photo 구조 내보내기
DROP TABLE IF EXISTS `db_pastoralstaff_photo`;
CREATE TABLE IF NOT EXISTS `db_pastoralstaff_photo` (
  `lifeno` varchar(50) NOT NULL,
  `photo` longtext DEFAULT NULL,
  PRIMARY KEY (`lifeno`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_pastoralwife 구조 내보내기
DROP TABLE IF EXISTS `db_pastoralwife`;
CREATE TABLE IF NOT EXISTS `db_pastoralwife` (
  `lifeno` varchar(20) NOT NULL COMMENT '생명번호',
  `nationality` varchar(20) NOT NULL COMMENT '국적',
  `name_ko` varchar(50) NOT NULL COMMENT '한글이름',
  `name_en` varchar(50) NOT NULL COMMENT '영문이름',
  `birthday` date NOT NULL COMMENT '생년월일',
  `phone` varchar(80) DEFAULT NULL COMMENT '전화번호',
  `home` varchar(200) DEFAULT NULL COMMENT '본가',
  `family` varchar(700) DEFAULT NULL COMMENT '가족사항',
  `health` text DEFAULT NULL COMMENT '건강사항',
  `other` text DEFAULT NULL COMMENT '기타사항',
  `lifeno_spouse` varchar(20) NOT NULL COMMENT '배우자 생명번호',
  `education` varchar(100) DEFAULT NULL COMMENT '학력',
  `suspend` tinyint(1) DEFAULT 0 COMMENT '논리삭제',
  `ovs_dept` int(3) NOT NULL DEFAULT 0 COMMENT '관리부서',
  PRIMARY KEY (`lifeno`),
  KEY `lifeno_spouse` (`lifeno_spouse`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_phone 구조 내보내기
DROP TABLE IF EXISTS `db_phone`;
CREATE TABLE IF NOT EXISTS `db_phone` (
  `church_sid` varchar(20) NOT NULL COMMENT '교회코드',
  `phone` varchar(50) DEFAULT NULL COMMENT '유선전화',
  `wmcphone` varchar(50) DEFAULT NULL COMMENT '인터넷전화',
  `address` varchar(1000) DEFAULT NULL,
  PRIMARY KEY (`church_sid`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_photoframe_binary 구조 내보내기
DROP TABLE IF EXISTS `db_photoframe_binary`;
CREATE TABLE IF NOT EXISTS `db_photoframe_binary` (
  `lifeno` varchar(50) NOT NULL,
  `photo` longblob DEFAULT NULL,
  PRIMARY KEY (`lifeno`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_position 구조 내보내기
DROP TABLE IF EXISTS `db_position`;
CREATE TABLE IF NOT EXISTS `db_position` (
  `position_cd` int(5) NOT NULL AUTO_INCREMENT,
  `lifeno` varchar(20) NOT NULL COMMENT '생명번호',
  `start_dt` date NOT NULL COMMENT '시작일',
  `end_dt` date NOT NULL DEFAULT '9999-12-31' COMMENT '종료일',
  `position` varchar(15) NOT NULL COMMENT '직책',
  PRIMARY KEY (`position_cd`),
  KEY `LifeNo` (`lifeno`) USING BTREE,
  KEY `Start_dt` (`start_dt`,`end_dt`) USING BTREE
) ENGINE=InnoDB AUTO_INCREMENT=7704 DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_position2 구조 내보내기
DROP TABLE IF EXISTS `db_position2`;
CREATE TABLE IF NOT EXISTS `db_position2` (
  `position2_cd` int(5) NOT NULL AUTO_INCREMENT,
  `lifeno` varchar(20) NOT NULL COMMENT '생명번호',
  `start_dt` date NOT NULL COMMENT '시작일',
  `end_dt` date NOT NULL DEFAULT '9999-12-31' COMMENT '종료일',
  `position2` varchar(15) NOT NULL COMMENT '특수직책',
  PRIMARY KEY (`position2_cd`),
  KEY `LifeNo` (`lifeno`) USING BTREE,
  KEY `Start_dt` (`start_dt`,`end_dt`) USING BTREE
) ENGINE=InnoDB AUTO_INCREMENT=528 DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_sermon 구조 내보내기
DROP TABLE IF EXISTS `db_sermon`;
CREATE TABLE IF NOT EXISTS `db_sermon` (
  `lifeno` varchar(20) NOT NULL COMMENT '생명번호',
  `score_avg` decimal(5,2) NOT NULL DEFAULT 0.00 COMMENT '발표점수',
  `subject_count` int(3) DEFAULT 0 COMMENT '발표개수',
  PRIMARY KEY (`lifeno`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_theological 구조 내보내기
DROP TABLE IF EXISTS `db_theological`;
CREATE TABLE IF NOT EXISTS `db_theological` (
  `theological_cd` int(5) NOT NULL AUTO_INCREMENT,
  `LifeNo` varchar(20) NOT NULL COMMENT '생명번호',
  `Level` varchar(10) NOT NULL COMMENT '예비생도 단계',
  `Start_dt` date NOT NULL COMMENT '시작일',
  `End_dt` date NOT NULL DEFAULT '9999-12-31' COMMENT '종료일',
  `Resign_dt` date DEFAULT NULL COMMENT '성도복귀일',
  `church_sid` varchar(15) NOT NULL COMMENT '추천교회',
  PRIMARY KEY (`theological_cd`),
  KEY `LifeNo` (`LifeNo`),
  KEY `Start_dt` (`Start_dt`,`End_dt`)
) ENGINE=InnoDB AUTO_INCREMENT=6177 DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_time_different 구조 내보내기
DROP TABLE IF EXISTS `db_time_different`;
CREATE TABLE IF NOT EXISTS `db_time_different` (
  `country` varchar(100) NOT NULL COMMENT '국가',
  `time_different` time NOT NULL COMMENT '시차',
  PRIMARY KEY (`country`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_title 구조 내보내기
DROP TABLE IF EXISTS `db_title`;
CREATE TABLE IF NOT EXISTS `db_title` (
  `title_cd` int(5) NOT NULL AUTO_INCREMENT,
  `lifeno` varchar(20) NOT NULL COMMENT '생명번호',
  `start_dt` date NOT NULL COMMENT '시작일',
  `end_dt` date NOT NULL DEFAULT '9999-12-31' COMMENT '종료일',
  `title` varchar(50) NOT NULL COMMENT '직분',
  `title_ordinary_date` date DEFAULT '1900-01-01' COMMENT '직분안수일',
  PRIMARY KEY (`title_cd`),
  KEY `LifeNo` (`lifeno`) USING BTREE,
  KEY `Start_dt` (`start_dt`,`end_dt`) USING BTREE
) ENGINE=InnoDB AUTO_INCREMENT=4534 DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_transfer 구조 내보내기
DROP TABLE IF EXISTS `db_transfer`;
CREATE TABLE IF NOT EXISTS `db_transfer` (
  `transfer_cd` int(5) NOT NULL AUTO_INCREMENT,
  `lifeno` varchar(20) NOT NULL COMMENT '생명번호',
  `start_dt` date NOT NULL COMMENT '시작일',
  `end_dt` date NOT NULL DEFAULT '9999-12-31' COMMENT '종료일',
  `church_sid` varchar(15) NOT NULL COMMENT '교회',
  PRIMARY KEY (`transfer_cd`),
  KEY `lifeno` (`lifeno`),
  KEY `start_dt` (`start_dt`,`end_dt`),
  KEY `church_sid` (`church_sid`)
) ENGINE=InnoDB AUTO_INCREMENT=11400 DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_union 구조 내보내기
DROP TABLE IF EXISTS `db_union`;
CREATE TABLE IF NOT EXISTS `db_union` (
  `union_cd` int(5) NOT NULL AUTO_INCREMENT,
  `church_sid_custom` varchar(20) NOT NULL COMMENT '교회코드(커스텀)',
  `start_dt` date DEFAULT NULL COMMENT '시작일',
  `end_dt` date DEFAULT NULL COMMENT '종료일',
  `union` int(5) NOT NULL COMMENT '연합회',
  PRIMARY KEY (`union_cd`),
  KEY `church_sid_custom` (`church_sid_custom`),
  KEY `start_dt` (`start_dt`,`end_dt`),
  KEY `union` (`union`)
) ENGINE=InnoDB AUTO_INCREMENT=694 DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_visa 구조 내보내기
DROP TABLE IF EXISTS `db_visa`;
CREATE TABLE IF NOT EXISTS `db_visa` (
  `visa_cd` int(5) NOT NULL AUTO_INCREMENT COMMENT '비자코드',
  `lifeno` varchar(20) NOT NULL COMMENT '생명번호',
  `Start_dt` date NOT NULL COMMENT '시작일',
  `End_dt` date NOT NULL DEFAULT '9999-12-31' COMMENT '종료일',
  `visa` varchar(50) NOT NULL COMMENT '비자종류',
  `memo` varchar(100) DEFAULT NULL COMMENT '메모',
  PRIMARY KEY (`visa_cd`),
  KEY `lifeno` (`lifeno`),
  KEY `Start_dt` (`Start_dt`,`End_dt`)
) ENGINE=InnoDB AUTO_INCREMENT=1806 DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.db_visa_photo 구조 내보내기
DROP TABLE IF EXISTS `db_visa_photo`;
CREATE TABLE IF NOT EXISTS `db_visa_photo` (
  `visa_cd` varchar(50) NOT NULL DEFAULT '',
  `photo` longtext DEFAULT NULL,
  PRIMARY KEY (`visa_cd`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 뷰 op_system.report_division_by_country 구조 내보내기
DROP VIEW IF EXISTS `report_division_by_country`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `report_division_by_country` (
	`dept_nm` VARCHAR(30) NOT NULL COMMENT '부서명' COLLATE 'utf8_general_ci',
	`user_nm` VARCHAR(20) NOT NULL COMMENT '엑셀의 사용자 이름으로 사용' COLLATE 'utf8_general_ci',
	`country_nm_ko` VARCHAR(50) NOT NULL COMMENT '국가명(한글)' COLLATE 'utf8_general_ci'
) ENGINE=MyISAM;

-- 프로시저 op_system.Routine_atten_detail 구조 내보내기
DROP PROCEDURE IF EXISTS `Routine_atten_detail`;
DELIMITER //
CREATE PROCEDURE `Routine_atten_detail`(
	IN `search_dt` DATE,
	IN `search_church` VARCHAR(50),
	IN `user_id` INT
)
BEGIN

	-- TRUNCATE op_system.temp_atten_detail;
	DELETE attenDetail FROM op_system.temp_atten_detail attenDetail WHERE attenDetail.user_id = user_id;
	
	INSERT INTO op_system.temp_atten_detail
	SELECT
	    esta.church_sid_custom AS church_sid_custom
	    ,churchlist.church_sid AS church_sid
	    ,churchlist.church_nm AS church_nm
	    ,churchlist.church_gb AS church_gb
	    ,esta.start_dt AS church_start_dt
	    ,esta.end_dt AS church_end_dt
	    ,IF(overseer.lifeno IS NULL, bleader.lifeno, overseer.lifeno) AS lifeno
	    ,IF(overseer.name_title IS NULL, bleader.name_title, overseer.name_title) AS name_title
	    ,IF(overseer.birthday IS NULL, bleader.birthday, overseer.birthday) AS birthday
	    ,IF(overseer.title IS NULL, bleader.title, overseer.title) AS title
	    ,IF(overseer.position IS NULL, bleader.position, overseer.position) AS posi
	    ,IF(overseer.nationality IS NULL, bleader.nationality, overseer.nationality) AS nationality
	    ,IF(overseer.appo_ovs IS NULL, bleader.appo_ovs, overseer.appo_ovs) AS appo_ovs
	    ,IF(overseer.trans_start_dt IS NULL, bleader.trans_start_dt, overseer.trans_start_dt) AS trans_start_dt
	    ,bleader.bleader_Start_dt AS bleader_start_dt
	    ,wife.lifeno AS lifeno_spouse
	    ,IF(title_spouse.Title IS NULL, wife.name_ko, CONCAT(wife.name_ko,'(',LEFT(title_spouse.Title,1),')')) AS name_title_spouse
	    ,wife.birthday AS birthday_spouse
	    ,title_spouse.Title AS title_spouse
	    ,position_spouse.position_Spouse AS position_spouse
	    ,atten.attendance_dt AS attendance_dt
	    ,atten.once_all AS once_all
	    ,atten.forth_all AS forth_all
	    ,atten.once_stu AS once_stu
	    ,atten.forth_stu AS forth_stu
	    ,atten.tithe_all AS tithe_all
	    ,atten.tithe_stu AS tithe_stu
	    ,atten.baptism_all AS baptism_all
	    ,atten.evangelist AS evangelist
	    ,atten.gl AS gl
	    ,atten.ul AS ul
	    ,user_id AS user_id
	    ,geo.country_nm_ko AS 'country'
	FROM op_system.db_history_church_establish AS esta
	LEFT JOIN op_system.db_churchlist_custom churchlist
	    ON esta.church_sid = churchlist.church_sid
	LEFT JOIN (
	    SELECT
	        trans.church_sid
	        ,pstaff.lifeno
	        ,IF(title.title IS NULL, pstaff.name_ko, CONCAT(pstaff.name_ko,'(',LEFT(title.title,1),')')) AS name_title
	        ,pstaff.birthday
	        ,title.Title
	        ,posi.Position
	        ,pstaff.nationality
	        ,pstaff.appo_ovs
	        ,trans.start_dt AS trans_start_dt
	    FROM op_system.db_pastoralstaff pstaff
	    INNER JOIN op_system.db_transfer trans
	        ON pstaff.lifeno = trans.LifeNo
	            AND LAST_DAY(search_dt) BETWEEN trans.Start_dt AND trans.End_dt
	    INNER JOIN op_system.db_position posi
	        ON pstaff.lifeno = posi.LifeNo
	            AND LAST_DAY(search_dt) BETWEEN posi.Start_dt AND posi.End_dt
	            AND posi.Position LIKE '당%'
	    LEFT JOIN op_system.db_title title
	        ON pstaff.lifeno = title.LifeNo
	            AND LAST_DAY(search_dt) BETWEEN title.Start_dt AND title.End_dt
	) AS overseer
	    ON esta.church_sid = overseer.church_sid
	LEFT JOIN (
	    SELECT
	        bleader.church_sid
	        ,pstaff.lifeno
	        ,IF(title.title IS NULL, pstaff.name_ko, CONCAT(pstaff.name_ko,'(',LEFT(title.title,1),')')) AS name_title
	        ,pstaff.birthday
	        ,title.Title
	        ,posi.Position
	        ,pstaff.nationality
	        ,trans.start_dt AS trans_start_dt
	        ,pstaff.appo_ovs
	        ,bleader.Start_dt AS bleader_start_dt
	    FROM op_system.db_pastoralstaff pstaff
	    INNER JOIN op_system.db_branchleader bleader
	        ON pstaff.lifeno = bleader.lifeno
	            AND LAST_DAY(search_dt) BETWEEN bleader.Start_dt AND bleader.End_dt
	            AND bleader.responsibility = '관리자'
	    LEFT JOIN op_system.db_position posi
	        ON pstaff.lifeno = posi.LifeNo
	            AND LAST_DAY(search_dt) BETWEEN posi.Start_dt AND posi.End_dt
	    LEFT JOIN op_system.db_position2 posi2
	        ON pstaff.lifeno = posi2.lifeno
	            AND LAST_DAY(search_dt) BETWEEN posi2.start_dt AND posi2.end_dt
	    LEFT JOIN op_system.db_theological theo
	        ON pstaff.lifeno = theo.LifeNo
	            AND LAST_DAY(search_dt) BETWEEN theo.Start_dt AND theo.End_dt
	    LEFT JOIN op_system.db_title title
	        ON pstaff.lifeno = title.LifeNo
	            AND LAST_DAY(search_dt) BETWEEN title.Start_dt AND title.End_dt
	    LEFT JOIN op_system.db_transfer trans
	        ON pstaff.lifeno = trans.lifeno
	            AND LAST_DAY(search_dt) BETWEEN trans.start_dt AND trans.end_dt
	    WHERE bleader.church_sid IS NOT NULL 
	        AND (posi.Position IS NOT NULL OR posi2.position2 IS NOT NULL OR theo.`Level` IS NOT NULL)
	) AS bleader
	    ON bleader.church_sid = esta.church_sid
	LEFT JOIN (
	    SELECT
	        esta.church_sid_custom
	        ,atten.church_sid
	        ,atten.attendance_dt
	        ,MAX(atten.once_all) AS once_all
	        ,MAX(atten.forth_all) AS forth_all
	        ,MAX(atten.once_stu) AS once_stu
	        ,MAX(atten.forth_stu) AS forth_stu
	        ,MAX(atten.tithe_all) AS tithe_all
	        ,MAX(atten.tithe_stu) AS tithe_stu
	        ,MAX(atten.baptism_all) AS baptism_all
	        ,MAX(atten.evangelist) AS evangelist
	        ,MAX(atten.gl) AS gl
	        ,MAX(atten.ul) AS ul
	    FROM op_system.db_history_church_establish esta
	    INNER JOIN op_system.db_attendance atten
	        ON esta.church_sid = atten.church_sid
	    WHERE atten.attendance_dt >= ADDDATE(search_dt, INTERVAL -10 YEAR) -- 엑셀 리소스 부족으로 최근 5년치만 가져옴
	    GROUP BY esta.church_sid_custom, atten.attendance_dt
	) atten
	    ON esta.church_sid_custom = atten.church_sid_custom
	LEFT JOIN op_system.db_pastoralwife wife
	    ON overseer.lifeno = wife.lifeno_spouse
	        OR bleader.lifeno = wife.lifeno_spouse
	LEFT JOIN op_system.db_title title_spouse
	    ON wife.lifeno = title_spouse.LifeNo
	        AND LAST_DAY(search_dt) BETWEEN title_spouse.Start_dt AND title_spouse.End_dt
	LEFT JOIN op_system.a_position_spouse position_spouse
	    ON overseer.position = position_spouse.position
	        OR bleader.position = position_spouse.position
	INNER JOIN op_system.db_geodata geo
	    ON churchlist.geo_cd = geo.geo_cd
	WHERE 
	    (LAST_DAY(search_dt) BETWEEN esta.start_dt AND esta.end_dt)
	    AND (churchlist.church_sid = search_church OR churchlist.main_church_cd = search_church)
	ORDER BY churchlist.sort_order, atten.attendance_dt ASC;
	
	INSERT INTO op_system.temp_atten_detail
	SELECT
	    esta.church_sid_custom
	    ,churchlist.church_sid
	    ,churchlist.church_nm
	    ,churchlist.church_gb
	    ,esta.start_dt AS church_start_dt
	    ,esta.end_dt AS church_end_dt
	    ,overseer.lifeno AS lifeno
	    ,overseer.name_title AS name_title
	    ,overseer.birthday AS birthday
	    ,overseer.title AS title
	    ,overseer.position AS posi
	    ,overseer.nationality AS nationality
	    ,overseer.appo_ovs AS appo_ovs
	    ,overseer.trans_start_dt AS trans_start_dt
	    ,NULL 
	    ,wife.lifeno
	    ,IF(title_spouse.Title IS NULL, wife.name_ko, CONCAT(wife.name_ko,'(',LEFT(title_spouse.Title,1),')')) AS name_title_spouse
	    ,wife.birthday AS birthday_spouse
	    ,title_spouse.Title AS title_spouse
	    ,position_spouse.position_Spouse AS position_spouse
	    ,atten.attendance_dt AS attendance_dt
	    ,atten.once_all AS once_all
	    ,atten.forth_all AS forth_all
	    ,atten.once_stu AS once_stu
	    ,atten.forth_stu AS forth_stu
	    ,atten.tithe_all AS tithe_all
	    ,atten.tithe_stu AS tithe_stu
	    ,atten.baptism_all AS baptism_all
	    ,atten.evangelist AS evangelist
	    ,atten.gl AS gl
	    ,atten.ul AS ul
	    ,user_id AS user_id
	    ,geo.country_nm_ko AS 'country'
	FROM op_system.db_history_church_establish esta
	LEFT JOIN op_system.db_churchlist_custom churchlist
	    ON REPLACE(esta.church_sid, 'MC', 'MM') = churchlist.church_sid
	LEFT JOIN (
	    SELECT
	        trans.church_sid
	        ,pstaff.lifeno
	        ,IF(title.title IS NULL, pstaff.name_ko, CONCAT(pstaff.name_ko,'(',LEFT(title.title,1),')')) AS name_title
	        ,pstaff.birthday
	        ,title.Title
	        ,posi.Position
	        ,pstaff.nationality
	        ,pstaff.appo_ovs
	        ,trans.start_dt AS trans_start_dt
	    FROM op_system.db_pastoralstaff pstaff
	    INNER JOIN op_system.db_transfer trans
	        ON pstaff.lifeno = trans.LifeNo
	            AND LAST_DAY(search_dt) BETWEEN trans.Start_dt AND trans.End_dt
	    INNER JOIN op_system.db_position posi
	        ON pstaff.lifeno = posi.LifeNo
	            AND LAST_DAY(search_dt) BETWEEN posi.Start_dt AND posi.End_dt
	            AND posi.Position LIKE '%당%'
	    LEFT JOIN op_system.db_title title
	        ON pstaff.lifeno = title.LifeNo
	            AND LAST_DAY(search_dt) BETWEEN title.Start_dt AND title.End_dt
	) AS overseer
	    ON esta.church_sid = overseer.church_sid
	LEFT JOIN (
	    SELECT
	        esta.church_sid_custom
	        ,atten.church_sid
	        ,atten.attendance_dt
	        ,MAX(atten.once_all) AS once_all
	        ,MAX(atten.forth_all) AS forth_all
	        ,MAX(atten.once_stu) AS once_stu
	        ,MAX(atten.forth_stu) AS forth_stu
	        ,MAX(atten.tithe_all) AS tithe_all
	        ,MAX(atten.tithe_stu) AS tithe_stu
	        ,MAX(atten.baptism_all) AS baptism_all
	        ,MAX(atten.evangelist) AS evangelist
	        ,MAX(atten.gl) AS gl
	        ,MAX(atten.ul) AS ul
	    FROM op_system.db_history_church_establish esta
	    INNER JOIN op_system.db_attendance atten
	        ON atten.church_sid = REPLACE(esta.church_sid, 'MC', 'MM')
	    WHERE atten.attendance_dt >= ADDDATE(search_dt, INTERVAL -10 YEAR) -- 엑셀 리소스 부족으로 최근 5년치만 가져옴
		 	  AND esta.church_sid = search_church
	    GROUP BY esta.church_sid_custom, atten.attendance_dt
	) atten
	    ON esta.church_sid_custom = atten.church_sid_custom
	LEFT JOIN op_system.db_pastoralwife wife
	    ON overseer.lifeno = wife.lifeno_spouse
	LEFT JOIN op_system.db_title title_spouse
	    ON wife.lifeno = title_spouse.LifeNo
	        AND LAST_DAY(search_dt) BETWEEN title_spouse.Start_dt AND title_spouse.End_dt
	LEFT JOIN op_system.a_position_spouse position_spouse
	    ON overseer.position = position_spouse.position
	INNER JOIN op_system.db_geodata geo
	    ON churchlist.geo_cd = geo.geo_cd
	WHERE (LAST_DAY(search_dt) BETWEEN esta.start_dt AND esta.end_dt)
	    AND (churchlist.church_sid = REPLACE(search_church, 'MC', 'MM'))
	ORDER BY churchlist.sort_order, atten.attendance_dt ASC;

END//
DELIMITER ;

-- 프로시저 op_system.Routine_atten_detail_main 구조 내보내기
DROP PROCEDURE IF EXISTS `Routine_atten_detail_main`;
DELIMITER //
CREATE PROCEDURE `Routine_atten_detail_main`(
	IN `search_dt` DATE,
	IN `user_id` INT
)
BEGIN
	
--	TRUNCATE op_system.temp_atten_detail;
	DELETE attenDetail FROM op_system.temp_atten_detail_main attenDetail WHERE attenDetail.user_id = user_id;
	
	INSERT INTO op_system.temp_atten_detail_main
	SELECT
		esta.church_sid_custom AS church_sid_custom
		,churchlist.church_sid AS church_sid
		,churchlist.church_nm AS church_nm
		,churchlist.church_gb AS church_gb
		,esta.start_dt AS church_start_dt
		,esta.end_dt AS church_end_dt
		,overseer.lifeno AS lifeno
		,overseer.name_title AS name_title
		,overseer.birthday AS birthday
		,overseer.title AS title
		,overseer.position AS posi
		,overseer.nationality AS nationality
		,overseer.appo_ovs AS appo_ovs
		,overseer.trans_start_dt AS trans_start_dt
		,NULL
		,wife.lifeno AS lifeno_spouse
		,IF(title_spouse.Title IS NULL, wife.name_ko, CONCAT(wife.name_ko,'(',LEFT(title_spouse.Title,1),')')) AS name_title_spouse
		,wife.birthday AS birthday_spouse
		,title_spouse.Title AS title_spouse
		,position_spouse.position_Spouse AS position_spouse
		,atten.attendance_dt AS attendance_dt
		,atten.once_all AS once_all
		,atten.forth_all AS forth_all
		,atten.once_stu AS once_stu
		,atten.forth_stu AS forth_stu
		,atten.tithe_all AS tithe_all
		,atten.tithe_stu AS tithe_stu
		,atten.baptism_all AS baptism_all
		,atten.evangelist AS evangelist
		,atten.gl AS gl
		,atten.ul AS ul
		,user_id AS user_id
		,geo.country_nm_ko AS 'country'
	FROM op_system.db_history_church_establish AS esta
	LEFT JOIN op_system.db_churchlist churchlist
		ON esta.church_sid = churchlist.church_sid
	LEFT JOIN 
		(
			SELECT
				trans.church_sid
				,pstaff.lifeno
				,IF(title.title IS NULL, pstaff.name_ko, CONCAT(pstaff.name_ko,'(',LEFT(title.title,1),')')) AS name_title
				,pstaff.birthday
				,title.Title
				,posi.Position
				,pstaff.nationality
				,pstaff.appo_ovs
				,trans.start_dt AS trans_start_dt
			FROM op_system.db_pastoralstaff pstaff
			LEFT JOIN op_system.db_transfer trans
				ON pstaff.lifeno = trans.LifeNo
					AND LAST_DAY(search_dt) BETWEEN trans.Start_dt AND trans.End_dt
			LEFT JOIN op_system.db_position posi
				ON pstaff.lifeno = posi.LifeNo
					AND LAST_DAY(search_dt) BETWEEN posi.Start_dt AND posi.End_dt
					AND posi.Position LIKE '%당%'
			LEFT JOIN op_system.db_title title
				ON pstaff.lifeno = title.LifeNo
					AND LAST_DAY(search_dt) BETWEEN title.Start_dt AND title.End_dt
			WHERE trans.church_sid IS NOT NULL
				AND posi.Position IS NOT NULL
		) AS overseer
		ON esta.church_sid = overseer.church_sid
	LEFT JOIN 
		(
			SELECT
				esta.church_sid_custom
				,atten.church_sid
				,atten.attendance_dt
				,MAX(atten.once_all) AS once_all
				,MAX(atten.forth_all) AS forth_all
				,MAX(atten.once_stu) AS once_stu
				,MAX(atten.forth_stu) AS forth_stu
				,MAX(atten.tithe_all) AS tithe_all
				,MAX(atten.tithe_stu) AS tithe_stu
				,MAX(atten.baptism_all) AS baptism_all
				,MAX(atten.evangelist) AS evangelist
				,MAX(atten.gl) AS gl
				,MAX(atten.ul) AS ul
			FROM op_system.db_history_church_establish esta
			LEFT JOIN op_system.db_attendance atten
				ON esta.church_sid = atten.church_sid
			WHERE atten.attendance_dt >= ADDDATE(search_dt, INTERVAL -3 YEAR) -- 엑셀 리소스 부족으로 최근 3년치만 가져옴
--				AND esta.church_sid = search_church
			GROUP BY esta.church_sid_custom, atten.attendance_dt
		) atten
		ON esta.church_sid_custom = atten.church_sid_custom
	LEFT JOIN op_system.db_pastoralwife wife
		ON overseer.lifeno = wife.lifeno_spouse
	LEFT JOIN op_system.db_title title_spouse
		ON wife.lifeno = title_spouse.LifeNo
			AND LAST_DAY(search_dt) BETWEEN title_spouse.Start_dt AND title_spouse.End_dt
	LEFT JOIN op_system.a_position_spouse position_spouse
		ON overseer.position = position_spouse.position
	LEFT JOIN common.users users
		ON users.user_dept = churchlist.ovs_dept
	LEFT JOIN op_system.db_geodata geo
		ON churchlist.geo_cd = geo.geo_cd
	WHERE 
		(LAST_DAY(search_dt) BETWEEN esta.start_dt AND esta.end_dt)
		AND (churchlist.church_gb = 'MC' OR churchlist.church_gb LIKE '%HBC%')
		AND (users.user_id = user_id)
	ORDER BY churchlist.sort_order, atten.attendance_dt ASC;
END//
DELIMITER ;

-- 프로시저 op_system.Routine_churchlist_by_time 구조 내보내기
DROP PROCEDURE IF EXISTS `Routine_churchlist_by_time`;
DELIMITER //
CREATE PROCEDURE `Routine_churchlist_by_time`(
	IN `search_dt` DATE



,
	IN `ovs_dept` INT


)
BEGIN
	TRUNCATE op_system.temp_churchlist_by_time;
	
	INSERT INTO op_system.temp_churchlist_by_time
	SELECT a.* 
		FROM op_system.db_churchlist a 
		LEFT JOIN op_system.db_history_church_establish b
			ON a.church_sid = b.church_sid
		WHERE LAST_DAY(search_dt) BETWEEN b.start_dt AND b.end_dt 
			AND IF(a.ovs_dept='',0,a.ovs_dept) = ovs_dept
		ORDER BY a.sort_order;
END//
DELIMITER ;

-- 프로시저 op_system.Routine_make_churchlist_custom 구조 내보내기
DROP PROCEDURE IF EXISTS `Routine_make_churchlist_custom`;
DELIMITER //
CREATE PROCEDURE `Routine_make_churchlist_custom`()
BEGIN

	SET @max_sort = (SELECT MAX(sort_order) FROM op_system.db_churchlist);

	TRUNCATE op_system.db_churchlist_custom;
	
	INSERT INTO op_system.db_churchlist_custom
	(SELECT 
		* 
		FROM op_system.db_churchlist
	UNION
	SELECT
		concat('MM',right(a.church_sid,CHARACTER_LENGTH(a.church_sid)-2))
		,concat(a.church_nm,' 본교회')
		,'MM' AS church_gb
		,a.manager_cd
		,a.main_church_cd
		,null
		,null
		,a.ovs_dept
		,a.suspend
		,a.sort_order + 1 AS sort_order
		,a.geo_cd
		FROM op_system.db_churchlist a WHERE church_sid LIKE 'MC%'
	ORDER BY sort_order);
	
	/*총회 추가*/
	INSERT 
		INTO op_system.db_churchlist_custom
		VALUES('HO001','총회','MC','','',null,null,'',0,@max_sort, NULL);
	
	/*엘로힘 연수원 추가*/
	INSERT 
		INTO op_system.db_churchlist_custom
		VALUES('IN001','엘로힘연수원','MC','','',null,null,'',0,@max_sort + 100, NULL);
		
	/*고앤컴 연수원 추가*/
	INSERT 
		INTO op_system.db_churchlist_custom
		VALUES('IN002','고앤컴연수원','MC','','',null,null,'',0,@max_sort + 200, NULL);
		
	/*전의산 연수원 추가*/
	INSERT 
		INTO op_system.db_churchlist_custom
		VALUES('IN003','전의산연수원','MC','','',null,null,'',0,@max_sort + 300, NULL);
		
	/*동백 연수원 추가*/
	INSERT 
		INTO op_system.db_churchlist_custom
		VALUES('IN004','동백연수원','MC','','',null,null,'',0,@max_sort + 400, NULL);
	
	/*제주 연수원 추가*/
	INSERT 
		INTO op_system.db_churchlist_custom
		VALUES('IN005','제주연수원','MC','','',null,null,'',0,@max_sort + 500, NULL);
		
	/*페루 영상물품실*/
	INSERT 
		INTO op_system.db_churchlist_custom
		VALUES('PS001','페루영상물품실','PS','','',null,NULL,'2',0,@max_sort + 600, NULL);
		
	/*db_churchlist에도 연수원 추가*/
	INSERT 
		INTO op_system.db_churchlist
		VALUES('IN001','엘로힘연수원','MC','','',null,null,'',0,@max_sort + 100, NULL);
		
	INSERT 
		INTO op_system.db_churchlist
		VALUES('IN002','고앤컴연수원','MC','','',null,null,'',0,@max_sort + 200, NULL);
		
	INSERT 
		INTO op_system.db_churchlist
		VALUES('IN003','전의산연수원','MC','','',null,null,'',0,@max_sort + 300, NULL);
		
	INSERT 
		INTO op_system.db_churchlist
		VALUES('IN004','동백연수원','MC','','',null,null,'',0,@max_sort + 400, NULL);
		
	INSERT 
		INTO op_system.db_churchlist
		VALUES('IN005','제주연수원','MC','','',null,null,'',0,@max_sort + 500, NULL);
END//
DELIMITER ;

-- 프로시저 op_system.Routine_pstaff_by_time 구조 내보내기
DROP PROCEDURE IF EXISTS `Routine_pstaff_by_time`;
DELIMITER //
CREATE PROCEDURE `Routine_pstaff_by_time`(
	IN `Search_dt` DATE,
	IN `department` INT
)
BEGIN
	
	SET @Max_dt = (SELECT MAX(a.attendance_dt) FROM op_system.db_attendance a);

	TRUNCATE op_system.temp_pstaff_by_time;
	
	INSERT INTO op_system.temp_pstaff_by_time 
	WITH basic AS (
		SELECT
			churchlist.church_nm '교회명'
			,church_admin.church_nm_en '영문교회명'
			,if(isnull(branchNM.church_nm),churchlist.church_nm,branchNM.church_nm) '지교회명'
			,if(isnull(branch_admin.church_nm_en),church_admin.church_nm_en,branch_admin.church_nm_en) '영문지교회명'
			,IFNULL(geoBranch.country_nm_ko, geo.country_nm_ko) '선교국가'
			,pstaff.lifeno '생명번호'
			,concat(pstaff.name_ko,ifnull(concat('(',left(title.Title,1),')'),'')) '한글이름(직분)'
			,pstaff.name_en '영문이름'
			,ifnull(if(`position`.`Position` like '%관리자%' or `position`.`Position` like '%당%' or `position`.`Position` like '%동%',`position`.`Position`,`theological`.`Level`),IF(`position2`.`position2` is not null,if(`position`.`position` is null,'직책없음',`position`.`position`),'직책없음')) '직책'
			,position2.position2 '직책2'
			,pstaff.birthday '생년월일'
			,pstaff.nationality '국적'
--			,IF(
--				IFNULL(appoint.start_dt,'1900-01-01')>IFNULL(pstaff.appo_ovs,'1900-01-01'),
--				appoint.start_dt,
--				pstaff.appo_ovs
--			) '(해외)최초발령일' -- 한국인: 해외발령일, 현지인: 최초발령일
			,IF(
	  	      pstaff.appo_ovs IS NULL, appoint.start_dt, 
		      pstaff.appo_ovs
			) '(해외)최초발령일' -- 한국인: 해외발령일, 현지인: 최초발령일
			,IF(`position`.`position` IN ( '당회장', '당회장대리','동역' ),
	              CASE 
						  WHEN `position`.`position` = '동역'
		              THEN IF(`position`.`start_dt` >= `belong`.`start_dt`, `position`.`start_dt`, `belong`.`start_dt`)
		              ELSE `belong`.`start_dt`
	              END
	              , NULL) AS '현당회발령일'
			,branchlist.Start_dt '관리자선임일'
			,IF(pstaff.wedding_dt is NULL,spouse.lifeno,if(LAST_DAY(Search_dt)<pstaff.wedding_dt,NULL,spouse.lifeno)) '배우자생번'
			,IF(pstaff.wedding_dt is NULL,concat(spouse.name_ko,ifnull(concat('(',left(title_spouse.Title,1),')'),'')),if(LAST_DAY(Search_dt)<pstaff.wedding_dt,'',concat(spouse.name_ko,ifnull(concat('(',left(title_spouse.Title,1),')'),'')))) '사모한글이름(직분)'
			,IF(pstaff.wedding_dt is NULL,spouse.name_en,if(LAST_DAY(Search_dt)<pstaff.wedding_dt,'',spouse.name_en)) '사모영문이름'
			,IF(pstaff.wedding_dt is NULL,if(spouse.lifeno is not null, spouseposition.position_Spouse, null),if(LAST_DAY(Search_dt)<pstaff.wedding_dt,'',if(spouse.lifeno is not null, spouseposition.position_Spouse, null))) as '사모직책'
			,IF(pstaff.wedding_dt is NULL,spouse.birthday,if(LAST_DAY(Search_dt)<pstaff.wedding_dt,NULL,spouse.birthday)) '배우자 생년월일'
			,if((ifnull(position.position,theological.level) like '당%' or ifnull(position.position,theological.level) like '동%'),null,theological.`Level`) '생도기수'
			,title.Title '직분'
			,pstaff.baptism '침례권'
			,unionnm.union_nm '연합회'
			,atten.once_all '전체1회'
			,atten.once_stu '학생1회'
			,atten.forth_all '전체4회'
			,atten.forth_stu '학생4회'
			,atten_main.once_all '본전체1회'
			,atten_main.once_stu '본학생1회'
			,atten_main.forth_all '본전체4회'
			,atten_main.forth_stu '본학생4회'
			,attenBranch.once_all '지전체1회'
			,attenBranch.once_stu '지학생1회'
			,attenBranch.forth_all '지전체4회'
			,attenBranch.forth_stu '지학생4회'
			,atten.tithe_stu '학생반차'
			,atten_main.tithe_stu '본학생반차'
			,attenBranch.tithe_stu '지학생반차'
			,pstaff.ovs_dept '관리부서'
			,churchlist.church_sid '교회코드'
			,IFNULL(branchNM.church_gb,churchlist.church_gb) '교회구분'
			,position2.Start_dt '직책2시작일'
			,theological.Start_dt '생도단계시작일'
			,branchlist.responsibility '소속구분'
			,pstaff.salary '유급'
			,visa.visa '선지자비자'
			,visa2.visa '배우자비자'
			,branchNM.church_sid '지교회코드'
		FROM op_system.db_pastoralstaff pstaff
			LEFT JOIN op_system.db_title title ON pstaff.lifeno = title.LifeNo AND (LAST_DAY(Search_dt) BETWEEN title.Start_dt AND title.End_dt) -- 직분테이블 결합
			LEFT JOIN op_system.db_position position ON pstaff.lifeno = position.LifeNo AND (LAST_DAY(Search_dt) BETWEEN position.Start_dt AND position.End_dt) -- 직책테이블 결합
			LEFT JOIN op_system.db_position2 position2 ON pstaff.lifeno = position2.LifeNo AND (LAST_DAY(Search_dt) BETWEEN position2.Start_dt AND position2.End_dt) -- 직책2테이블 결합
			LEFT JOIN op_system.db_transfer belong ON pstaff.lifeno = belong.lifeno AND (LAST_DAY(Search_dt) BETWEEN belong.Start_dt AND belong.End_dt) -- 발령이력 결합
			LEFT JOIN op_system.db_churchlist churchlist ON belong.church_sid = churchlist.church_sid -- 본교회명 결합
			LEFT JOIN op_system.db_history_church_establish churchesta ON churchlist.church_sid = churchesta.church_sid -- 설립이력 결합
			LEFT JOIN op_system.db_union uniondb 
				ON churchesta.church_sid_custom = uniondb.church_sid_custom
					AND LAST_DAY(Search_dt) BETWEEN uniondb.start_dt AND uniondb.end_dt -- 연합회코드 결합
			LEFT JOIN op_system.a_union unionnm ON uniondb.`union` = unionnm.union_cd -- 연합회명 결합
			LEFT JOIN op_system.db_attendance atten 
				ON churchlist.church_sid = atten.church_sid 
					AND atten.attendance_dt = IF(LAST_DAY(Search_dt) > @Max_dt,@Max_dt,DATE_ADD(DATE_ADD(LAST_DAY(Search_dt),INTERVAL 1 DAY),INTERVAL -1 MONTH)) -- AND atten.attendance_dt = DATE_ADD(DATE_ADD(LAST_DAY(Search_dt),INTERVAL 1 DAY),INTERVAL -1 MONTH) -- 본교회출석 결합
			LEFT JOIN op_system.db_attendance atten_main
				ON REPLACE(churchlist.church_sid,'MC','MM') = atten_main.church_sid 
					AND atten_main.attendance_dt = IF(LAST_DAY(Search_dt) > @Max_dt,@Max_dt,DATE_ADD(DATE_ADD(LAST_DAY(Search_dt),INTERVAL 1 DAY),INTERVAL -1 MONTH)) -- AND atten.attendance_dt = DATE_ADD(DATE_ADD(LAST_DAY(Search_dt),INTERVAL 1 DAY),INTERVAL -1 MONTH) -- 본교회출석 결합
			LEFT JOIN op_system.db_branchleader branchlist ON pstaff.lifeno = branchlist.lifeno 
				AND (LAST_DAY(Search_dt) BETWEEN branchlist.Start_dt AND branchlist.End_dt)
--				AND branchlist.responsibility LIKE '%관리자%' -- 지교회코드 결합		
			LEFT JOIN op_system.db_churchlist branchNM ON branchlist.church_sid = branchNM.church_sid -- 지교회명 결합
			LEFT JOIN op_system.db_attendance attenBranch 
				ON branchlist.church_sid = attenBranch.church_sid 
					AND attenBranch.attendance_dt = IF(LAST_DAY(Search_dt) > @Max_dt,@Max_dt,DATE_ADD(DATE_ADD(LAST_DAY(Search_dt),INTERVAL 1 DAY),INTERVAL -1 MONTH)) -- AND attenBranch.attendance_dt = LAST_DAY(Search_dt) -- 지교회출석 결합
			LEFT JOIN op_system.db_theological theological ON pstaff.LifeNo = theological.LifeNo AND (LAST_DAY(Search_dt) BETWEEN theological.Start_dt AND theological.End_dt) -- 생도기수 결합
			LEFT JOIN op_system.db_pastoralwife spouse ON pstaff.lifeno = spouse.lifeno_spouse -- 배우자 정보 결합
			LEFT JOIN op_system.a_position_spouse spouseposition on ifnull(position.position,theological.level) = spouseposition.position -- 배우자 직책 결합
			LEFT JOIN (SELECT op_system.db_position.lifeno,min(op_system.db_position.Start_dt) start_dt,op_system.db_position.Position FROM op_system.db_position
							WHERE
							(op_system.db_position.Position IN ('당회장','당회장대리','동역'))
							GROUP BY
							op_system.db_position.LifeNo) as appoint
				on appoint.lifeno = pstaff.lifeno -- 최초발령일 결합
			LEFT JOIN (SELECT * FROM op_system.db_title) title_spouse on title_spouse.lifeno = spouse.lifeno AND (LAST_DAY(Search_dt) BETWEEN title_spouse.Start_dt AND title_spouse.End_dt) -- 선지자 직책에 따른 배우자 직책 결합
			LEFT JOIN op_system.a_churchlist_admin church_admin on churchlist.church_sid = church_admin.church_sid
			LEFT JOIN op_system.a_churchlist_admin branch_admin on branchlist.church_sid = branch_admin.church_sid
			LEFT JOIN op_system.db_geodata geo ON churchlist.geo_cd = geo.geo_cd
			LEFT JOIN op_system.db_geodata geoBranch ON branchNM.geo_cd = geoBranch.geo_cd
			LEFT JOIN op_system.db_visa visa ON visa.lifeno = pstaff.lifeno AND (LAST_DAY(Search_dt) BETWEEN visa.Start_dt AND visa.End_dt) -- 선지자 비자정보
			LEFT JOIN op_system.db_visa visa2 ON visa2.lifeno = spouse.lifeno AND (LAST_DAY(Search_dt) BETWEEN visa2.Start_dt AND visa2.End_dt) -- 배우자 비자정보
		WHERE (churchlist.church_nm is not null 
					AND IFNULL(IF(INSTR(if(isnull(branchNM.church_nm),churchlist.church_nm,branchNM.church_nm),' ')>0,LEFT(if(isnull(branchNM.church_nm),churchlist.church_nm,branchNM.church_nm),INSTR(if(isnull(branchNM.church_nm),churchlist.church_nm,branchNM.church_nm),' ')-1),if(isnull(branchNM.church_nm),churchlist.church_nm,branchNM.church_nm)),church_admin.country) is not null 
					AND not (if(`position`.`Position` like '%관리자%' or `position`.`Position` like '%당%' or `position`.`Position` like '%동%',`position`.`Position`,`theological`.`Level`) is null and position2.position2 is null))
				AND NOT (position2.position2 is null AND if(`position`.`Position` like '%관리자%' or `position`.`Position` like '%당%' or `position`.`Position` like '%동%',`position`.`Position`,`theological`.`Level`) LIKE '%역장%')
	--			AND (branchNM.main_church_cd = belong.church_sid OR branchNM.church_nm IS NULL)
				AND pstaff.ovs_dept = department
		ORDER BY `교회명`,`(해외)최초발령일`
	)
	SELECT * FROM basic
	
	UNION 
	
	SELECT
		main.church_nm
		,main_admin.church_nm_en
		,branch.church_nm
		,branch_admin.church_nm_en
		,NULL
		,NULL
		,NULL
		,NULL
		,NULL
		,NULL
		,NULL
		,NULL
		,NULL
		,NULL
		,NULL
		,NULL
		,NULL
		,NULL
		,NULL
		,NULL
		,NULL
		,NULL
		,NULL
		,NULL
		,main_atten.once_all
		,main_atten.forth_all
		,main_atten.once_stu
		,main_atten.forth_stu
		,NULL
		,NULL
		,NULL
		,NULL
		,branch_atten.once_all
		,branch_atten.forth_all
		,branch_atten.once_stu
		,branch_atten.forth_stu
		,main_atten.tithe_stu
		,NULL
		,branch_atten.tithe_stu
		,branch.ovs_dept
		,main.church_sid
		,branch.church_gb
		,NULL
		,NULL
		,NULL
		,NULL
		,NULL 
		,NULL 
		,branch.church_sid
	FROM op_system.db_churchlist branch
	LEFT JOIN op_system.db_history_church_establish esta
		ON esta.church_sid = branch.church_sid AND (LAST_DAY(search_dt) BETWEEN esta.start_dt AND esta.end_dt)
	LEFT JOIN op_system.db_churchlist main
		ON main.church_sid = branch.main_church_cd
	LEFT JOIN op_system.a_churchlist_admin main_admin
		ON main_admin.church_sid = main.church_sid
	LEFT JOIN op_system.a_branch_admin branch_admin
		ON branch_admin.church_sid = branch.church_sid
	LEFT JOIN op_system.db_attendance branch_atten
		ON branch_atten.church_sid = branch.church_sid
			AND branch_atten.attendance_dt = IF(LAST_DAY(Search_dt) > @Max_dt,@Max_dt,DATE_ADD(DATE_ADD(LAST_DAY(Search_dt),INTERVAL 1 DAY),INTERVAL -1 MONTH))
	LEFT JOIN op_system.db_attendance main_atten
		ON main_atten.church_sid = main.church_sid
			AND main_atten.attendance_dt = IF(LAST_DAY(Search_dt) > @Max_dt,@Max_dt,DATE_ADD(DATE_ADD(LAST_DAY(Search_dt),INTERVAL 1 DAY),INTERVAL -1 MONTH))
	WHERE branch.church_nm NOT IN (SELECT `지교회명` FROM basic) 
		AND branch.suspend = 0 
		AND branch.ovs_dept = department;
	
END//
DELIMITER ;

-- 프로시저 op_system.Routine_statistic_by_church 구조 내보내기
DROP PROCEDURE IF EXISTS `Routine_statistic_by_church`;
DELIMITER //
CREATE PROCEDURE `Routine_statistic_by_church`(
	IN `search_dt` DATE
)
    COMMENT '교회통계 임시테이블 작성루틴'
BEGIN
	TRUNCATE op_system.temp_statistic_by_church;
	
	INSERT INTO op_system.temp_statistic_by_church
	SELECT
--			g.country '국가'
			IF(INSTR(a1.church_nm,' ')=0,IF(g.country is null,a1.church_nm,g.country),LEFT(a1.church_nm,INSTR(a1.church_nm,' ')-1)) '국가'
			,d.union_nm '연합회'
			,a1.church_nm '관리교회'
--			,IFNULL(f.`한글이름(직분)`,f1.`한글이름(직분)`) '관리자명'
			,f.`한글이름(직분)` '관리자명'
			,a.`LBC`
			,a.`LPBC`
			,a.`소계`
			,b.once_all '전체1회'
			,b.forth_all '전체4회'
			,b.once_stu '학생1회'
			,b.forth_stu '학생4회'
			,b.tithe_stu '학생반차'
			,b.baptism_all '침례'
			,b.evangelist '전도인'
			,b.gl '지역장'
			,b.ul '구역장'
			,c.`인원`
			,c.`당회장`
			,c.`당회장대리`
			,c.`동역`
			,c.`예비생도`
--			,IF(DATEDIFF(CURDATE(),search_dt)<90,e.`지교회관리자`,c.`지교회관리자`)
--			,IF(DATEDIFF(CURDATE(),search_dt)<90,e.`예배소관리자`,c.`예배소관리자`)
			,c.`지교회관리자`
			,c.`예배소관리자`
			,c.`당사모`
			,c.`당대리사모`
			,c.`동사모`
			,c.`생도사모`
			,c.`지관자사모`
			,c.`예관자사모`
			,a1.ovs_dept '관리부서'
			,a1.sort_order '정렬순서'
		FROM op_system.temp_churchlist_by_time a1
		LEFT JOIN
			(SELECT 
					DISTINCT b.church_nm 'church_nm'
					,COUNT(if(a.church_gb<>'BC',NULL,a.church_gb)) 'LBC'
					,COUNT(if(a.church_gb<>'PBC',NULL,a.church_gb)) 'LPBC'
					,COUNT(if(a.church_gb<>'BC',NULL,a.church_gb)) + COUNT(if(a.church_gb<>'PBC',NULL,a.church_gb)) '소계'
				FROM op_system.temp_churchlist_by_time a
				LEFT JOIN op_system.temp_churchlist_by_time b
					ON a.main_church_cd = b.church_sid
				WHERE b.church_nm IS NOT NULL
				GROUP BY a.main_church_cd) a -- 교회개수 통계
			ON a1.church_nm = a.church_nm
		LEFT JOIN 
			(SELECT 
					b.church_nm
					,a.*
				FROM op_system.db_attendance a 
				LEFT JOIN op_system.db_churchlist_custom b
					ON a.church_sid = b.church_sid
				WHERE a.attendance_dt = search_dt AND b.church_gb IN ('MC','HBC')
				ORDER BY b.sort_order) b -- 출석인원
			ON a1.church_nm = b.church_nm
		LEFT JOIN
			(SELECT a.`교회명`
					,COUNT(if(a.`직책`<>'당회장',NULL,a.`직책`))+COUNT(if(a.`직책`<>'당회장대리',NULL,a.`직책`))+COUNT(if(a.`직책`<>'동역',NULL,a.`직책`)) '인원'
					,COUNT(if(a.`직책`<>'당회장',NULL,a.`직책`)) '당회장'
					,COUNT(if(a.`직책`<>'당회장대리',NULL,a.`직책`)) '당회장대리'
					,COUNT(if(a.`직책`<>'동역',NULL,a.`직책`)) '동역'
					,COUNT(if(a.`생도기수` NOT LIKE '%생도%',NULL,a.`생도기수`)) '예비생도'
					,COUNT(if(a.`직책` NOT LIKE '%지교회%',NULL,a.`직책`)) '지교회관리자'
					,COUNT(if(a.`직책` NOT LIKE '%예배소%',NULL,a.`직책`)) '예배소관리자'
					,COUNT(if(a.`사모직책`<>'당사모',NULL,a.`직책`)) '당사모'
					,COUNT(if(a.`사모직책`<>'당대리사모',NULL,a.`직책`)) '당대리사모'
					,COUNT(if(a.`사모직책`<>'동사모',NULL,a.`직책`)) '동사모'
					,COUNT(if(a.`사모직책`<>'생도사모',NULL,a.`직책`)) '생도사모'
					,COUNT(if(a.`사모직책`<>'지관자사모',NULL,a.`직책`)) '지관자사모'
					,COUNT(if(a.`사모직책`<>'예관자사모',NULL,a.`직책`)) '예관자사모'
				FROM op_system.temp_pstaff_by_time a 
				GROUP BY a.`교회명`) c
			ON a1.church_nm = c.`교회명`
		LEFT JOIN
			(SELECT 
					d.church_nm
					,b.union_nm
					,b.sort_order
				FROM op_system.db_union a 
				LEFT JOIN op_system.a_union b
					ON a.`union` = b.union_cd
				LEFT JOIN op_system.db_history_church_establish c
					ON a.church_sid_custom = c.church_sid_custom AND (LAST_DAY(search_dt) BETWEEN c.start_dt AND c.end_dt)
				LEFT JOIN op_system.db_churchlist_custom d
					ON c.church_sid = d.church_sid
				WHERE LAST_DAY(search_dt) BETWEEN a.start_dt AND a.end_dt) d
			ON a1.church_nm = d.church_nm
		LEFT JOIN
			(SELECT
					a.main_church 'church_nm'
					,COUNT(if(a.position='지교회관리자',a.position,NULL)) '지교회관리자'
					,COUNT(if(a.position='예배소관리자',a.position,NULL)) '예배소관리자'
				FROM op_system.a_branch_admin a
				WHERE NOT a.main_church LIKE '%old%'
				GROUP BY a.main_church) e
			ON a1.church_nm = e.church_nm
		LEFT JOIN op_system.temp_pstaff_by_time f
			ON a1.church_sid = f.`교회코드`
				-- MC는 당회장, HBC는 관리자명 뜨도록 조정 / HBC는 관리 지교회, 예배소가 없다는 전제가 있음
				AND IF(a1.church_gb='MC',f.`직책` LIKE '%당%',IFNULL(f.`직책` LIKE '%지교회%', f.`직책` LIKE '%예배소%'))
--		LEFT JOIN op_system.temp_pstaff_by_time f1
--			ON a1.church_sid = f1.`교회코드`
--				AND f1.`직책` LIKE '%지교회관리자%' AND a1.church_gb = 'HBC'
		LEFT JOIN op_system.a_churchlist_admin g
			ON a1.church_sid = g.church_sid
		WHERE a1.church_gb IN ('MC','HBC')
		ORDER BY a1.sort_order;
END//
DELIMITER ;

-- 프로시저 op_system.Routine_statistic_by_church_all 구조 내보내기
DROP PROCEDURE IF EXISTS `Routine_statistic_by_church_all`;
DELIMITER //
CREATE PROCEDURE `Routine_statistic_by_church_all`(
	IN `search_dt` DATE
)
    COMMENT '전체교회통계 임시테이블 작성루틴'
BEGIN
	
	TRUNCATE op_system.temp_statistic_by_church_all;
	
	INSERT INTO op_system.temp_statistic_by_church_all
	WITH cte AS
(
   SELECT
		churchlist.church_nm '교회명'
		,churchlist.church_gb '교회형태'
		,esta.start_dt '설립일'
		,IFNULL(POSITION1.position,POSITION2.Position) '직책'
		,CONCAT(pstaff.name_ko,IF(title.Title IS NULL,'',CONCAT('(',LEFT(title.Title,1),')'))) '관리자'
		,atten.once_all '전체1회'
		,atten.forth_all '전체4회'
		,atten.once_stu '학생1회'
		,atten.forth_stu '학생4회'
		,atten.tithe_stu '학생반차'
		,atten.baptism_all '전체침례'
		,atten.evangelist '전도인'
		,atten.gl '지역장'
		,atten.ul '구역장'
		,churchlist.ovs_dept '관리부서'
		,union_nm.union_nm '연합회'
		,IF(mainchurch.church_nm IS NULL AND churchlist.church_gb='MM',REPLACE(churchlist.church_nm,' 본교회',''),mainchurch.church_nm) '본교회명'
		,churchlist.sort_order '정렬순서'
		,NULL '교회개수'
	FROM 
	(
		SELECT 
			* 
			FROM op_system.temp_churchlist_by_time
		UNION
		SELECT
			REPLACE(a.church_sid,'MC','MM')
			,concat(a.church_nm,' 본교회')
			,'MM' AS church_gb
			,a.manager_cd
			,a.main_church_cd
			,null
			,null
			,a.ovs_dept
			,a.suspend
			,a.sort_order + 1 AS sort_order
			,a.geo_cd
			FROM op_system.temp_churchlist_by_time a WHERE church_gb = 'MC'
		ORDER BY sort_order
	) churchlist
	LEFT JOIN op_system.db_churchlist mainchurch
		ON mainchurch.church_sid=churchlist.main_church_cd
	LEFT JOIN op_system.db_attendance atten
		ON churchlist.church_sid=atten.church_sid AND atten.attendance_dt = search_dt
	LEFT JOIN op_system.db_transfer trans
		ON trans.church_sid=REPLACE(churchlist.church_sid,'MM','MC') AND LAST_DAY(search_dt) BETWEEN trans.start_dt AND trans.end_dt
	LEFT JOIN op_system.db_branchleader branchleader
		ON branchleader.church_sid=churchlist.church_sid 
			AND branchleader.responsibility LIKE '%관리자%'
			AND LAST_DAY(search_dt) BETWEEN branchleader.Start_dt AND branchleader.End_dt
	LEFT JOIN op_system.db_pastoralstaff pstaff
		ON (pstaff.lifeno=trans.lifeno OR pstaff.lifeno=branchleader.lifeno)
	LEFT JOIN op_system.db_title title
		ON (title.LifeNo=pstaff.lifeno OR title.LifeNo=branchleader.lifeno) AND LAST_DAY(search_dt) BETWEEN title.Start_dt AND title.End_dt
	LEFT JOIN op_system.db_position POSITION1
		ON POSITION1.LifeNo=pstaff.lifeno AND (LAST_DAY(search_dt) BETWEEN POSITION1.Start_dt AND POSITION1.End_dt)
	LEFT JOIN op_system.db_position POSITION2	
		ON POSITION2.LifeNo=branchleader.lifeno AND LAST_DAY(search_dt) BETWEEN POSITION2.Start_dt AND POSITION2.End_dt
	LEFT JOIN op_system.db_history_church_establish esta
		ON esta.church_sid=REPLACE(churchlist.church_sid,'MM','MC')
	LEFT JOIN op_system.db_union union_id
		ON union_id.church_sid_custom=esta.church_sid_custom AND LAST_DAY(search_dt) BETWEEN union_id.start_dt AND union_id.end_dt
	LEFT JOIN op_system.a_union union_nm
		ON union_nm.union_cd=union_id.`union`
)
	
	SELECT
		*
	FROM cte a
	WHERE NOT (`교회형태` IN ('MC','MM','HBC') AND `직책` NOT LIKE '%당%')
		OR (`교회형태` IN ('HBC') AND `직책` IS NULL) -- 'HBC에 관리자가 아무도 없을 때 표시하기 위해  필요함
		OR (`교회형태` IN ('HBC') AND `직책` LIKE '%관리자%') -- 'HBC'에 지교회(예배소) 관리자가 있는 경우 표시하기 위해 필요함
	GROUP BY `교회명`
		-- GROUP BY는 교회명마다 하나씩만 나오도록 하기 위한 설정

	UNION
	
	SELECT 
		DISTINCT a.`교회명`
		,a.`교회형태`
		,a.`설립일`
		,NULL `직책`
		,NULL `관리자`
		,a.`전체1회`
		,a.`전체4회`
		,a.`학생1회`
		,a.`학생4회`
		,a.`학생반차`
		,a.`전체침례`
		,a.`전도인`
		,a.`지역장`
		,a.`구역장`
		,a.`관리부서`
		,a.`연합회`
		,a.`본교회명`
		,a.`정렬순서`
		,a.`교회개수`
	FROM cte a
	WHERE (`교회형태` IN ('MC','MM','HBC') AND `직책` NOT LIKE '%당%')
		AND `교회명` NOT IN 
			(
				SELECT `교회명` FROM cte 
				WHERE NOT (`교회형태` IN ('MC','MM','HBC') AND `직책` NOT LIKE '%당%')
					OR (`교회형태` IN ('HBC') AND `직책` LIKE '%관리자%') 
					-- HBC에 지교회관리자가 아닌 사람이 소속되어 있을 경우 오류가 나기 때문에 설정필요함.
					-- 이렇게 하면 지관자가 없는 HBC는 목록에서 표시 안되게 되어 있음.
			)
--	GROUP BY `교회명`
	
	ORDER BY `정렬순서`;
END//
DELIMITER ;

-- 프로시저 op_system.Routine_statistic_by_Country 구조 내보내기
DROP PROCEDURE IF EXISTS `Routine_statistic_by_Country`;
DELIMITER //
CREATE PROCEDURE `Routine_statistic_by_Country`(
	IN `Search_dt` DATE,
	IN `department` INT
)
    COMMENT '국가별 통계자료 임시테이블 작성'
BEGIN
	TRUNCATE op_system.temp_statistic_by_country;
	
	INSERT INTO op_SYSTEM.temp_statistic_by_country
	SELECT 
		geo.country_nm_ko AS '국가'
		,count(if(a.`교회구분`='MM',a.`교회구분`,NULL))+count(if(a.`교회구분`='BC',a.`교회구분`,NULL))+count(if(a.`교회구분`='PBC',a.`교회구분`,NULL)) '소계'
		,count(if(a.`교회구분` IN ('MM','HBC'),a.`교회구분`,NULL)) MC
		,count(if(a.`교회구분`='BC',a.`교회구분`,NULL)) BC
		,count(if(a.`교회구분`='PBC',a.`교회구분`,NULL)) PBC
		,sum(if(a.`교회구분` IN ('MC','HBC'),NULL,c.once_all)) '전체1회'
		,sum(if(a.`교회구분` IN ('MC','HBC'),NULL,c.forth_all)) '전체4회'
		,sum(if(a.`교회구분` IN ('MC','HBC'),NULL,c.once_stu)) '학생1회'
		,sum(if(a.`교회구분` IN ('MC','HBC'),NULL,c.forth_stu)) '학생4회'
		,sum(if(a.`교회구분` IN ('MC','HBC'),NULL,c.tithe_stu)) '반차'
		,sum(if(a.`교회구분` IN ('MC','HBC'),NULL,c.baptism_all)) '침례'
		,sum(if(a.`교회구분` IN ('MC','HBC'),NULL,c.evangelist)) '전도인'
		,sum(if(a.`교회구분` IN ('MC','HBC'),NULL,c.gl)) '지역장'
		,sum(if(a.`교회구분` IN ('MC','HBC'),NULL,c.ul)) '구역장'
		,IF(ISNULL(e.`인원`),0,e.`인원`)
		,IF(ISNULL(e.`당회장`),0,e.`당회장`)
		,IF(ISNULL(e.`당회장대리`),0,e.`당회장대리`)
		,IF(ISNULL(e.`동역`),0,e.`동역`)
		,IF(ISNULL(e.`생도`),0,e.`생도`)
		,IF(ISNULL(e.`지교회관리자`),0,e.`지교회관리자`)
		,IF(ISNULL(e.`예배소관리자`),0,e.`예배소관리자`)
		,IF(ISNULL(e.`당사모`),0,e.`당사모`)
		,IF(ISNULL(e.`당대리사모`),0,e.`당대리사모`)
		,IF(ISNULL(e.`동사모`),0,e.`동사모`)
		,IF(ISNULL(e.`생도사모`),0,e.`생도사모`)
		,IF(ISNULL(e.`지관자사모`),0,e.`지관자사모`)
		,IF(ISNULL(e.`예관자사모`),0,e.`예관자사모`)
		,a.`관리부서`
		FROM 
			(SELECT 
				CEsta.church_sid '교회코드'
				,CList.church_nm '한글교회명'
				,if(CList.church_gb IN ('MC','HBC'),'MM',CList.church_gb) '교회구분'
				,CList.sort_order '정렬순서'
				,Clist.ovs_dept '관리부서'
				,Clist.geo_cd
				FROM op_system.temp_churchlist_by_time CList
				LEFT JOIN op_system.db_history_church_establish CEsta
					ON CList.church_sid = CEsta.church_sid
				WHERE LAST_DAY(Search_dt) BETWEEN CEsta.start_dt AND CEsta.end_dt) a
		LEFT JOIN op_system.a_churchlist_admin b 
			ON a.`교회코드` = b.church_sid
		LEFT JOIN op_system.db_attendance c 
			ON REPLACE(a.`교회코드`,'MC','MM') = c.church_sid and c.attendance_dt = Search_dt
		LEFT JOIN op_system.db_geodata geo
			ON geo.geo_cd = a.geo_cd
		LEFT JOIN
			(SELECT 
					P_Staff.`선교국가` AS '선교국가'
					,COUNT(if(P_Staff.`직책`='당회장',P_Staff.`직책`,NULL))+COUNT(if(P_Staff.`직책`='당회장대리',P_Staff.`직책`,NULL))+COUNT(if(P_Staff.`직책`='동역',P_Staff.`직책`,NULL)) '인원'
					,COUNT(if(P_Staff.`직책`='당회장',P_Staff.`직책`,NULL)) '당회장'
					,COUNT(if(P_Staff.`직책`='당회장대리',P_Staff.`직책`,NULL)) '당회장대리'
					,COUNT(if(P_Staff.`직책`='동역',P_Staff.`직책`,NULL)) '동역'
					,COUNT(if(ISNULL(P_Staff.`생도기수`),NULL,P_Staff.`생도기수`)) '생도'
--					,BranchAdmin.`지교회관리자`
--					,BranchAdmin.`예배소관리자`
					,COUNT(if(P_Staff.`직책`='지교회관리자',P_Staff.`직책`,NULL)) '지교회관리자'
					,COUNT(if(P_Staff.`직책`='예배소관리자',P_Staff.`직책`,NULL)) '예배소관리자'
					,COUNT(if(P_Staff.`사모직책`='당사모',P_Staff.`사모직책`,NULL)) '당사모'
					,COUNT(if(P_Staff.`사모직책`='당대리사모',P_Staff.`사모직책`,NULL)) '당대리사모'
					,COUNT(if(P_Staff.`사모직책`='동사모',P_Staff.`사모직책`,NULL)) '동사모'
					,COUNT(if(P_Staff.`사모직책`='생도사모',P_Staff.`사모직책`,NULL)) '생도사모'
					,If((LAST_DAY(NOW() - interval 2 month) + interval 1 DAY) > search_dt,NULL,COUNT(if(P_Staff.`사모직책`='지관자사모',P_Staff.`사모직책`,NULL))) '지관자사모'
					,If((LAST_DAY(NOW() - interval 2 month) + interval 1 DAY) > search_dt,NULL,COUNT(if(P_Staff.`사모직책`='예관자사모',P_Staff.`사모직책`,NULL))) '예관자사모'
				FROM op_system.temp_pstaff_by_time P_staff
--				LEFT JOIN
--					(SELECT
--						badmin.country '선교국가'
--						,If((LAST_DAY(NOW() - interval 2 month) + interval 1 DAY) > search_dt,NULL,COUNT(if(badmin.position='지교회관리자',badmin.position,NULL))) '지교회관리자'
--						,If((LAST_DAY(NOW() - interval 2 month) + interval 1 DAY) > search_dt,NULL,COUNT(if(badmin.position='예배소관리자',badmin.position,NULL))) '예배소관리자'
--						FROM op_system.a_branch_admin Badmin
--						GROUP BY badmin.country) BranchAdmin
--					ON P_Staff.`선교국가`=BranchAdmin.`선교국가`
				GROUP BY P_Staff.`선교국가`) e
			ON geo.country_nm_ko = e.`선교국가`
--		WHERE REPLACE(IF(IF(INSTR(a.`한글교회명`,' ')>0,LEFT(a.`한글교회명`,INSTR(a.`한글교회명`,' ')-1),a.`한글교회명`) NOT IN (SELECT ctry_nm FROM op_system.db_country),b.country,IF(INSTR(a.`한글교회명`,' ')>0,LEFT(a.`한글교회명`,INSTR(a.`한글교회명`,' ')-1),a.`한글교회명`)),'제2','') IN (SELECT ctry_nm FROM OP_system.db_country)
		GROUP BY `국가`
		ORDER BY a.`정렬순서`;
END//
DELIMITER ;

-- 프로시저 op_system.Routine_statistic_by_pstaff 구조 내보내기
DROP PROCEDURE IF EXISTS `Routine_statistic_by_pstaff`;
DELIMITER //
CREATE PROCEDURE `Routine_statistic_by_pstaff`(
	IN `search_dt` DATE
)
BEGIN
	TRUNCATE op_system.temp_statistic_by_pstaff;
	
	INSERT INTO op_system.temp_statistic_by_pstaff
	SELECT
			geo.country_nm_ko '국가'
			,d.union_nm '연합회'
			,a1.church_nm '관리교회'
			,f.`한글이름(직분)` '관리자명'
			,b.once_all '전체1회'
			,c.`인원`
			,c.`당회장`
			,c.`당회장대리`
			,c.`동역`
			,c.`예비생도`
--			,IF(DATEDIFF(CURDATE(),search_dt)<90,e.`지교회관리자`,c.`지교회관리자`)
--			,IF(DATEDIFF(CURDATE(),search_dt)<90,e.`예배소관리자`,c.`예배소관리자`)
			,c.`지교회관리자`
			,c.`예배소관리자`
			,c.`침례권`
			,c.`목사`
			,c.`장로`
			,c.`전도사`
			,c.`집사`
			,c.`형제`
			,a1.ovs_dept '관리부서'
			,a1.sort_order '정렬순서'
		FROM op_system.temp_churchlist_by_time a1
		LEFT JOIN
			(SELECT 
					DISTINCT b.church_nm 'church_nm'
					,COUNT(if(a.church_gb<>'BC',NULL,a.church_gb)) 'LBC'
					,COUNT(if(a.church_gb<>'PBC',NULL,a.church_gb)) 'LPBC'
					,COUNT(if(a.church_gb<>'BC',NULL,a.church_gb)) + COUNT(IF(a.church_gb<>'PBC',NULL,a.church_gb)) '소계'
				FROM op_system.temp_churchlist_by_time a
				LEFT JOIN op_system.temp_churchlist_by_time b
					ON a.main_church_cd = b.church_sid
				WHERE b.church_nm IS NOT NULL
				GROUP BY a.main_church_cd) a -- 교회개수 통계
			ON a1.church_nm = a.church_nm
		LEFT JOIN 
			(SELECT 
					b.church_nm
					,a.*
				FROM op_system.db_attendance a 
				LEFT JOIN op_system.db_churchlist_custom b
					ON a.church_sid = b.church_sid
				WHERE a.attendance_dt = search_dt AND b.church_gb IN ('MC','HBC')
				ORDER BY b.sort_order) b -- 출석인원
			ON a1.church_nm = b.church_nm
		LEFT JOIN
			(SELECT a.`교회명`
					,COUNT(if(a.`직책`<>'당회장',NULL,a.`직책`))+COUNT(if(a.`직책`<>'당회장대리',NULL,a.`직책`))+COUNT(if(a.`직책`<>'동역',NULL,a.`직책`)) '인원'
					,COUNT(if(a.`직책`<>'당회장',NULL,a.`직책`)) '당회장'
					,COUNT(if(a.`직책`<>'당회장대리',NULL,a.`직책`)) '당회장대리'
					,COUNT(if(a.`직책`<>'동역',NULL,a.`직책`)) '동역'
					,COUNT(if(a.`생도기수` NOT LIKE '%생도%',NULL,a.`생도기수`)) '예비생도'
					,COUNT(if(a.`직책` NOT LIKE '%지교회%',NULL,a.`직책`)) '지교회관리자'
					,COUNT(if(a.`직책` NOT LIKE '%예배소%',NULL,a.`직책`)) '예배소관리자'
					,COUNT(if(a.`사모직책`<>'당사모',NULL,a.`직책`)) '당사모'
					,COUNT(if(a.`사모직책`<>'당대리사모',NULL,a.`직책`)) '당대리사모'
					,COUNT(if(a.`사모직책`<>'동사모',NULL,a.`직책`)) '동사모'
					,COUNT(if(a.`사모직책`<>'생도사모',NULL,a.`직책`)) '생도사모'
					,COUNT(if(a.`사모직책`<>'지관자사모',NULL,a.`직책`)) '지관자사모'
					,COUNT(if(a.`사모직책`<>'예관자사모',NULL,a.`직책`)) '예관자사모'
					,COUNT(if(a.`침례권`<>'유',NULL,a.`침례권`)) '침례권'
					,COUNT(if(a.`직분`<>'목사',NULL,a.`직분`)) '목사'
					,COUNT(if(a.`직분`<>'장로',NULL,a.`직분`)) '장로'
					,COUNT(if(a.`직분`<>'전도사',NULL,a.`직분`)) '전도사'
					,COUNT(if(a.`직분`<>'집사',NULL,a.`직분`)) '집사'
					,COUNT(if(a.`직분` IS NULL AND (a.`직책` LIKE '%당%' OR a.`직책` LIKE '%동%' OR a.`직책` LIKE '%관리자%'),'형제',NULL)) '형제'
				FROM op_system.temp_pstaff_by_time a 
				GROUP BY a.`교회명`) c
			ON a1.church_nm = c.`교회명`
		LEFT JOIN
			(SELECT 
					d.church_nm
					,b.union_nm
					,b.sort_order
				FROM op_system.db_union a 
				LEFT JOIN op_system.a_union b
					ON a.`union` = b.union_cd
				LEFT JOIN op_system.db_history_church_establish c
					ON a.church_sid_custom = c.church_sid_custom AND (LAST_DAY(search_dt) BETWEEN c.start_dt AND c.end_dt)
				LEFT JOIN op_system.db_churchlist_custom d
					ON c.church_sid = d.church_sid
				WHERE LAST_DAY(search_dt) BETWEEN a.start_dt AND a.end_dt) d
			ON a1.church_nm = d.church_nm
		LEFT JOIN
			(SELECT
					a.main_church 'church_nm'
					,COUNT(if(a.position='지교회관리자',a.position,NULL)) '지교회관리자'
					,COUNT(if(a.position='예배소관리자',a.position,NULL)) '예배소관리자'
				FROM op_system.a_branch_admin a
				WHERE NOT a.main_church LIKE '%old%'
				GROUP BY a.main_church) e
			ON a1.church_nm = e.church_nm
		LEFT JOIN op_system.temp_pstaff_by_time f
			ON a1.church_nm = f.`교회명` AND 
				-- MC는 당회장, HBC는 관리자명 뜨도록 조정 / HBC는 관리 지교회, 예배소가 없다는 전제가 있음
				IF(a1.church_gb='MC',f.`직책` LIKE '%당%',IFNULL(f.`직책` LIKE '%지교회%', f.`직책` LIKE '%예배소%'))
 		LEFT JOIN op_system.a_churchlist_admin g
			ON a1.church_sid = g.church_sid
		LEFT JOIN op_system.db_geodata geo
			ON a1.geo_cd = geo.geo_cd
		WHERE a1.church_gb IN ('MC','HBC')
		ORDER BY a1.sort_order;
END//
DELIMITER ;

-- 프로시저 op_system.Routine_test 구조 내보내기
DROP PROCEDURE IF EXISTS `Routine_test`;
DELIMITER //
CREATE PROCEDURE `Routine_test`(
	IN `search_dt` DATE,
	IN `search_church` VARCHAR(50),
	IN `user_id` INT
)
BEGIN

-- TRUNCATE op_system.temp_atten_detail;
DELETE attenDetail FROM op_system.temp_atten_detail attenDetail WHERE attenDetail.user_id = user_id;

INSERT INTO op_system.temp_atten_detail
SELECT
    esta.church_sid_custom AS church_sid_custom
    ,churchlist.church_sid AS church_sid
    ,churchlist.church_nm AS church_nm
    ,churchlist.church_gb AS church_gb
    ,esta.start_dt AS church_start_dt
    ,esta.end_dt AS church_end_dt
    ,IF(overseer.lifeno IS NULL, bleader.lifeno, overseer.lifeno) AS lifeno
    ,IF(overseer.name_title IS NULL, bleader.name_title, overseer.name_title) AS name_title
    ,IF(overseer.birthday IS NULL, bleader.birthday, overseer.birthday) AS birthday
    ,IF(overseer.title IS NULL, bleader.title, overseer.title) AS title
    ,IF(overseer.position IS NULL, bleader.position, overseer.position) AS posi
    ,IF(overseer.nationality IS NULL, bleader.nationality, overseer.nationality) AS nationality
    ,IF(overseer.appo_ovs IS NULL, bleader.appo_ovs, overseer.appo_ovs) AS appo_ovs
    ,IF(overseer.trans_start_dt IS NULL, bleader.trans_start_dt, overseer.trans_start_dt) AS trans_start_dt
    ,bleader.bleader_Start_dt AS bleader_start_dt
    ,wife.lifeno AS lifeno_spouse
    ,IF(title_spouse.Title IS NULL, wife.name_ko, CONCAT(wife.name_ko,'(',LEFT(title_spouse.Title,1),')')) AS name_title_spouse
    ,wife.birthday AS birthday_spouse
    ,title_spouse.Title AS title_spouse
    ,position_spouse.position_Spouse AS position_spouse
    ,atten.attendance_dt AS attendance_dt
    ,atten.once_all AS once_all
    ,atten.forth_all AS forth_all
    ,atten.once_stu AS once_stu
    ,atten.forth_stu AS forth_stu
    ,atten.tithe_all AS tithe_all
    ,atten.tithe_stu AS tithe_stu
    ,atten.baptism_all AS baptism_all
    ,atten.evangelist AS evangelist
    ,atten.gl AS gl
    ,atten.ul AS ul
    ,user_id AS user_id
    ,geo.country_nm_ko AS 'country'
FROM op_system.db_history_church_establish AS esta
LEFT JOIN op_system.db_churchlist_custom churchlist
    ON esta.church_sid = churchlist.church_sid
LEFT JOIN (
    SELECT
        trans.church_sid
        ,pstaff.lifeno
        ,IF(title.title IS NULL, pstaff.name_ko, CONCAT(pstaff.name_ko,'(',LEFT(title.title,1),')')) AS name_title
        ,pstaff.birthday
        ,title.Title
        ,posi.Position
        ,pstaff.nationality
        ,pstaff.appo_ovs
        ,trans.start_dt AS trans_start_dt
    FROM op_system.db_pastoralstaff pstaff
    INNER JOIN op_system.db_transfer trans
        ON pstaff.lifeno = trans.LifeNo
            AND LAST_DAY(search_dt) BETWEEN trans.Start_dt AND trans.End_dt
    INNER JOIN op_system.db_position posi
        ON pstaff.lifeno = posi.LifeNo
            AND LAST_DAY(search_dt) BETWEEN posi.Start_dt AND posi.End_dt
            AND posi.Position LIKE '당%'
    LEFT JOIN op_system.db_title title
        ON pstaff.lifeno = title.LifeNo
            AND LAST_DAY(search_dt) BETWEEN title.Start_dt AND title.End_dt
) AS overseer
    ON esta.church_sid = overseer.church_sid
LEFT JOIN (
    SELECT
        bleader.church_sid
        ,pstaff.lifeno
        ,IF(title.title IS NULL, pstaff.name_ko, CONCAT(pstaff.name_ko,'(',LEFT(title.title,1),')')) AS name_title
        ,pstaff.birthday
        ,title.Title
        ,posi.Position
        ,pstaff.nationality
        ,trans.start_dt AS trans_start_dt
        ,pstaff.appo_ovs
        ,bleader.Start_dt AS bleader_start_dt
    FROM op_system.db_pastoralstaff pstaff
    INNER JOIN op_system.db_branchleader bleader
        ON pstaff.lifeno = bleader.lifeno
            AND LAST_DAY(search_dt) BETWEEN bleader.Start_dt AND bleader.End_dt
            AND bleader.responsibility = '관리자'
    LEFT JOIN op_system.db_position posi
        ON pstaff.lifeno = posi.LifeNo
            AND LAST_DAY(search_dt) BETWEEN posi.Start_dt AND posi.End_dt
    LEFT JOIN op_system.db_position2 posi2
        ON pstaff.lifeno = posi2.lifeno
            AND LAST_DAY(search_dt) BETWEEN posi2.start_dt AND posi2.end_dt
    LEFT JOIN op_system.db_theological theo
        ON pstaff.lifeno = theo.LifeNo
            AND LAST_DAY(search_dt) BETWEEN theo.Start_dt AND theo.End_dt
    LEFT JOIN op_system.db_title title
        ON pstaff.lifeno = title.LifeNo
            AND LAST_DAY(search_dt) BETWEEN title.Start_dt AND title.End_dt
    LEFT JOIN op_system.db_transfer trans
        ON pstaff.lifeno = trans.lifeno
            AND LAST_DAY(search_dt) BETWEEN trans.start_dt AND trans.end_dt
    WHERE bleader.church_sid IS NOT NULL 
        AND (posi.Position IS NOT NULL OR posi2.position2 IS NOT NULL OR theo.`Level` IS NOT NULL)
) AS bleader
    ON bleader.church_sid = esta.church_sid
LEFT JOIN (
    SELECT
        esta.church_sid_custom
        ,atten.church_sid
        ,atten.attendance_dt
        ,MAX(atten.once_all) AS once_all
        ,MAX(atten.forth_all) AS forth_all
        ,MAX(atten.once_stu) AS once_stu
        ,MAX(atten.forth_stu) AS forth_stu
        ,MAX(atten.tithe_all) AS tithe_all
        ,MAX(atten.tithe_stu) AS tithe_stu
        ,MAX(atten.baptism_all) AS baptism_all
        ,MAX(atten.evangelist) AS evangelist
        ,MAX(atten.gl) AS gl
        ,MAX(atten.ul) AS ul
    FROM op_system.db_history_church_establish esta
    INNER JOIN op_system.db_attendance atten
        ON esta.church_sid = atten.church_sid
    WHERE atten.attendance_dt >= ADDDATE(search_dt, INTERVAL -10 YEAR) -- 엑셀 리소스 부족으로 최근 5년치만 가져옴
    GROUP BY esta.church_sid_custom, atten.attendance_dt
) atten
    ON esta.church_sid_custom = atten.church_sid_custom
LEFT JOIN op_system.db_pastoralwife wife
    ON overseer.lifeno = wife.lifeno_spouse
        OR bleader.lifeno = wife.lifeno_spouse
LEFT JOIN op_system.db_title title_spouse
    ON wife.lifeno = title_spouse.LifeNo
        AND LAST_DAY(search_dt) BETWEEN title_spouse.Start_dt AND title_spouse.End_dt
LEFT JOIN op_system.a_position_spouse position_spouse
    ON overseer.position = position_spouse.position
        OR bleader.position = position_spouse.position
INNER JOIN op_system.db_geodata geo
    ON churchlist.geo_cd = geo.geo_cd
WHERE 
    (LAST_DAY(search_dt) BETWEEN esta.start_dt AND esta.end_dt)
    AND (churchlist.church_sid = search_church OR churchlist.main_church_cd = search_church)
ORDER BY churchlist.sort_order, atten.attendance_dt ASC;

INSERT INTO op_system.temp_atten_detail
SELECT
    esta.church_sid_custom
    ,churchlist.church_sid
    ,churchlist.church_nm
    ,churchlist.church_gb
    ,esta.start_dt AS church_start_dt
    ,esta.end_dt AS church_end_dt
    ,overseer.lifeno AS lifeno
    ,overseer.name_title AS name_title
    ,overseer.birthday AS birthday
    ,overseer.title AS title
    ,overseer.position AS posi
    ,overseer.nationality AS nationality
    ,overseer.appo_ovs AS appo_ovs
    ,overseer.trans_start_dt AS trans_start_dt
    ,NULL 
    ,wife.lifeno
    ,IF(title_spouse.Title IS NULL, wife.name_ko, CONCAT(wife.name_ko,'(',LEFT(title_spouse.Title,1),')')) AS name_title_spouse
    ,wife.birthday AS birthday_spouse
    ,title_spouse.Title AS title_spouse
    ,position_spouse.position_Spouse AS position_spouse
    ,atten.attendance_dt AS attendance_dt
    ,atten.once_all AS once_all
    ,atten.forth_all AS forth_all
    ,atten.once_stu AS once_stu
    ,atten.forth_stu AS forth_stu
    ,atten.tithe_all AS tithe_all
    ,atten.tithe_stu AS tithe_stu
    ,atten.baptism_all AS baptism_all
    ,atten.evangelist AS evangelist
    ,atten.gl AS gl
    ,atten.ul AS ul
    ,user_id AS user_id
    ,geo.country_nm_ko AS 'country'
FROM op_system.db_history_church_establish esta
LEFT JOIN op_system.db_churchlist_custom churchlist
    ON REPLACE(esta.church_sid, 'MC', 'MM') = churchlist.church_sid
LEFT JOIN (
    SELECT
        trans.church_sid
        ,pstaff.lifeno
        ,IF(title.title IS NULL, pstaff.name_ko, CONCAT(pstaff.name_ko,'(',LEFT(title.title,1),')')) AS name_title
        ,pstaff.birthday
        ,title.Title
        ,posi.Position
        ,pstaff.nationality
        ,pstaff.appo_ovs
        ,trans.start_dt AS trans_start_dt
    FROM op_system.db_pastoralstaff pstaff
    INNER JOIN op_system.db_transfer trans
        ON pstaff.lifeno = trans.LifeNo
            AND LAST_DAY(search_dt) BETWEEN trans.Start_dt AND trans.End_dt
    INNER JOIN op_system.db_position posi
        ON pstaff.lifeno = posi.LifeNo
            AND LAST_DAY(search_dt) BETWEEN posi.Start_dt AND posi.End_dt
            AND posi.Position LIKE '%당%'
    LEFT JOIN op_system.db_title title
        ON pstaff.lifeno = title.LifeNo
            AND LAST_DAY(search_dt) BETWEEN title.Start_dt AND title.End_dt
) AS overseer
    ON esta.church_sid = overseer.church_sid
LEFT JOIN (
    SELECT
        esta.church_sid_custom
        ,atten.church_sid
        ,atten.attendance_dt
        ,MAX(atten.once_all) AS once_all
        ,MAX(atten.forth_all) AS forth_all
        ,MAX(atten.once_stu) AS once_stu
        ,MAX(atten.forth_stu) AS forth_stu
        ,MAX(atten.tithe_all) AS tithe_all
        ,MAX(atten.tithe_stu) AS tithe_stu
        ,MAX(atten.baptism_all) AS baptism_all
        ,MAX(atten.evangelist) AS evangelist
        ,MAX(atten.gl) AS gl
        ,MAX(atten.ul) AS ul
    FROM op_system.db_history_church_establish esta
    INNER JOIN op_system.db_attendance atten
        ON atten.church_sid = esta.church_sid
    WHERE atten.attendance_dt >= ADDDATE(search_dt, INTERVAL -10 YEAR) -- 엑셀 리소스 부족으로 최근 5년치만 가져옴
	 	  AND esta.church_sid = search_church
    GROUP BY esta.church_sid_custom, atten.attendance_dt
) atten
    ON esta.church_sid_custom = atten.church_sid_custom
LEFT JOIN op_system.db_pastoralwife wife
    ON overseer.lifeno = wife.lifeno_spouse
LEFT JOIN op_system.db_title title_spouse
    ON wife.lifeno = title_spouse.LifeNo
        AND LAST_DAY(search_dt) BETWEEN title_spouse.Start_dt AND title_spouse.End_dt
LEFT JOIN op_system.a_position_spouse position_spouse
    ON overseer.position = position_spouse.position
INNER JOIN op_system.db_geodata geo
    ON churchlist.geo_cd = geo.geo_cd
WHERE (LAST_DAY(search_dt) BETWEEN esta.start_dt AND esta.end_dt)
    AND (churchlist.church_sid = REPLACE(search_church, 'MC', 'MM'))
ORDER BY churchlist.sort_order, atten.attendance_dt ASC;

END//
DELIMITER ;

-- 테이블 op_system.temp_atten_detail 구조 내보내기
DROP TABLE IF EXISTS `temp_atten_detail`;
CREATE TABLE IF NOT EXISTS `temp_atten_detail` (
  `church_sid_custom` int(11) DEFAULT NULL,
  `church_sid` varchar(50) DEFAULT NULL,
  `church_nm` varchar(50) DEFAULT NULL,
  `church_gb` varchar(50) DEFAULT NULL,
  `church_start_dt` date DEFAULT NULL,
  `church_end_dt` date DEFAULT NULL,
  `lifeno` varchar(50) DEFAULT NULL,
  `name_title` varchar(50) DEFAULT NULL,
  `birthday` date DEFAULT NULL,
  `title` varchar(50) DEFAULT NULL,
  `position` varchar(50) DEFAULT NULL,
  `nationality` varchar(50) DEFAULT NULL,
  `appo_ovs` varchar(50) DEFAULT NULL,
  `trans_start_dt` date DEFAULT NULL,
  `bleader_start_dt` date DEFAULT NULL,
  `lifeno_spouse` varchar(50) DEFAULT NULL,
  `name_title_spouse` varchar(50) DEFAULT NULL,
  `birthday_spouse` date DEFAULT NULL,
  `title_spouse` varchar(50) DEFAULT NULL,
  `position_spouse` varchar(50) DEFAULT NULL,
  `attendance_dt` date DEFAULT NULL,
  `출석(전체 1회)` int(11) DEFAULT NULL,
  `출석(전체 4회)` int(11) DEFAULT NULL,
  `출석(학생이상 1회)` int(11) DEFAULT NULL,
  `출석(학생이상 4회)` int(11) DEFAULT NULL,
  `반차(전체)` int(11) DEFAULT NULL,
  `반차(학생이상)` int(11) DEFAULT NULL,
  `침례(전체)` int(11) DEFAULT NULL,
  `고정전도인` int(11) DEFAULT NULL,
  `지역장` int(11) DEFAULT NULL,
  `구역장` int(11) DEFAULT NULL,
  `user_id` int(11) NOT NULL,
  `country` varchar(50) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.temp_atten_detail_main 구조 내보내기
DROP TABLE IF EXISTS `temp_atten_detail_main`;
CREATE TABLE IF NOT EXISTS `temp_atten_detail_main` (
  `church_sid_custom` int(11) DEFAULT NULL,
  `church_sid` varchar(50) DEFAULT NULL,
  `church_nm` varchar(50) DEFAULT NULL,
  `church_gb` varchar(50) DEFAULT NULL,
  `church_start_dt` date DEFAULT NULL,
  `church_end_dt` date DEFAULT NULL,
  `lifeno` varchar(50) DEFAULT NULL,
  `name_title` varchar(50) DEFAULT NULL,
  `birthday` date DEFAULT NULL,
  `title` varchar(50) DEFAULT NULL,
  `position` varchar(50) DEFAULT NULL,
  `nationality` varchar(50) DEFAULT NULL,
  `appo_ovs` varchar(50) DEFAULT NULL,
  `trans_start_dt` date DEFAULT NULL,
  `bleader_start_dt` date DEFAULT NULL,
  `lifeno_spouse` varchar(50) DEFAULT NULL,
  `name_title_spouse` varchar(50) DEFAULT NULL,
  `birthday_spouse` date DEFAULT NULL,
  `title_spouse` varchar(50) DEFAULT NULL,
  `position_spouse` varchar(50) DEFAULT NULL,
  `attendance_dt` date DEFAULT NULL,
  `출석(전체 1회)` int(11) DEFAULT NULL,
  `출석(전체 4회)` int(11) DEFAULT NULL,
  `출석(학생이상 1회)` int(11) DEFAULT NULL,
  `출석(학생이상 4회)` int(11) DEFAULT NULL,
  `반차(전체)` int(11) DEFAULT NULL,
  `반차(학생이상)` int(11) DEFAULT NULL,
  `침례(전체)` int(11) DEFAULT NULL,
  `고정전도인` int(11) DEFAULT NULL,
  `지역장` int(11) DEFAULT NULL,
  `구역장` int(11) DEFAULT NULL,
  `user_id` int(11) NOT NULL,
  `country` varchar(50) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.temp_churchlist_by_time 구조 내보내기
DROP TABLE IF EXISTS `temp_churchlist_by_time`;
CREATE TABLE IF NOT EXISTS `temp_churchlist_by_time` (
  `church_sid` varchar(20) NOT NULL COMMENT '교회코드',
  `church_nm` varchar(50) NOT NULL COMMENT '교회명',
  `church_gb` varchar(15) NOT NULL COMMENT '교회형태',
  `manager_cd` varchar(100) DEFAULT NULL COMMENT '관리자 생명번호',
  `main_church_cd` varchar(20) DEFAULT NULL COMMENT '관리교회 코드',
  `start_dt` date DEFAULT NULL COMMENT '시작일',
  `end_dt` date DEFAULT NULL COMMENT '종료일',
  `ovs_dept` varchar(15) DEFAULT NULL COMMENT '해외국 관리부서',
  `suspend` tinyint(1) NOT NULL DEFAULT 0 COMMENT '논리삭제',
  `sort_order` int(10) NOT NULL COMMENT '정렬순서',
  `geo_cd` int(10) DEFAULT NULL,
  PRIMARY KEY (`church_sid`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.temp_pstaff_by_time 구조 내보내기
DROP TABLE IF EXISTS `temp_pstaff_by_time`;
CREATE TABLE IF NOT EXISTS `temp_pstaff_by_time` (
  `교회명` varchar(50) DEFAULT NULL,
  `영문교회명` varchar(100) DEFAULT NULL,
  `지교회명` varchar(50) DEFAULT NULL,
  `영문지교회명` varchar(100) DEFAULT NULL,
  `선교국가` varchar(50) DEFAULT NULL,
  `생명번호` varchar(20) DEFAULT NULL,
  `한글이름(직분)` varchar(50) DEFAULT NULL,
  `영문이름` varchar(50) DEFAULT NULL,
  `직책` varchar(20) DEFAULT NULL,
  `직책2` varchar(20) DEFAULT NULL,
  `생년월일` date DEFAULT NULL,
  `국적` varchar(20) DEFAULT NULL,
  `최초발령일` date DEFAULT NULL,
  `현당회발령일` date DEFAULT NULL,
  `관리자선임일` date DEFAULT NULL,
  `배우자생번` varchar(20) DEFAULT NULL,
  `사모한글이름(직분)` varchar(30) DEFAULT NULL,
  `사모영문이름` varchar(50) DEFAULT NULL,
  `사모직책` varchar(20) DEFAULT NULL,
  `배우자 생년월일` date DEFAULT NULL,
  `생도기수` varchar(20) DEFAULT NULL,
  `직분` varchar(20) DEFAULT NULL,
  `침례권` varchar(20) DEFAULT NULL,
  `연합회` varchar(20) DEFAULT NULL,
  `전체1회` int(5) DEFAULT NULL,
  `학생1회` int(5) DEFAULT NULL,
  `전체4회` int(5) DEFAULT NULL,
  `학생4회` int(5) DEFAULT NULL,
  `본전체1회` int(5) DEFAULT NULL,
  `본학생1회` int(5) DEFAULT NULL,
  `본전체4회` int(5) DEFAULT NULL,
  `본학생4회` int(5) DEFAULT NULL,
  `지전체1회` int(5) DEFAULT NULL,
  `지학생1회` int(5) DEFAULT NULL,
  `지전체4회` int(5) DEFAULT NULL,
  `지학생4회` int(5) DEFAULT NULL,
  `학생반차` int(5) DEFAULT NULL,
  `본학생반차` int(5) DEFAULT NULL,
  `지학생반차` int(5) DEFAULT NULL,
  `관리부서` int(3) NOT NULL DEFAULT 0 COMMENT '관리부서',
  `교회코드` varchar(20) DEFAULT NULL COMMENT '교회코드',
  `교회구분` varchar(15) DEFAULT NULL COMMENT '교회형태',
  `직책2시작일` date DEFAULT NULL,
  `생도단계시작일` date DEFAULT NULL,
  `소속구분` varchar(10) DEFAULT NULL,
  `유급` int(11) DEFAULT NULL,
  `선지자비자` varchar(20) DEFAULT NULL,
  `배우자비자` varchar(20) DEFAULT NULL,
  `지교회코드` varchar(20) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.temp_statistic_by_church 구조 내보내기
DROP TABLE IF EXISTS `temp_statistic_by_church`;
CREATE TABLE IF NOT EXISTS `temp_statistic_by_church` (
  `국가` varchar(15) NOT NULL,
  `연합회` varchar(15) DEFAULT NULL,
  `관리교회` varchar(50) NOT NULL,
  `관리자명` varchar(30) DEFAULT NULL,
  `LBC` int(4) DEFAULT NULL,
  `LPBC` int(4) DEFAULT NULL,
  `소계` int(5) DEFAULT NULL,
  `전체1회` int(10) DEFAULT NULL,
  `전체4회` int(10) DEFAULT NULL,
  `학생1회` int(10) DEFAULT NULL,
  `학생4회` int(10) DEFAULT NULL,
  `반차` int(7) DEFAULT NULL,
  `침례` int(5) DEFAULT NULL,
  `전도인` int(5) DEFAULT NULL,
  `지역장` int(5) DEFAULT NULL,
  `구역장` int(5) DEFAULT NULL,
  `인원` int(5) DEFAULT NULL,
  `당회장` int(5) DEFAULT NULL,
  `당회장대리` int(5) DEFAULT NULL,
  `동역` int(5) DEFAULT NULL,
  `생도` int(5) DEFAULT NULL,
  `지교회관리자` int(5) DEFAULT NULL,
  `예배소관리자` int(5) DEFAULT NULL,
  `당사모` int(5) DEFAULT NULL,
  `당대리사모` int(5) DEFAULT NULL,
  `동사모` int(5) DEFAULT NULL,
  `생도사모` int(5) DEFAULT NULL,
  `지관자사모` int(5) DEFAULT NULL,
  `예관자사모` int(5) DEFAULT NULL,
  `관리부서` int(3) NOT NULL DEFAULT 0 COMMENT '관리부서',
  `정렬순서` int(10) NOT NULL DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.temp_statistic_by_church_all 구조 내보내기
DROP TABLE IF EXISTS `temp_statistic_by_church_all`;
CREATE TABLE IF NOT EXISTS `temp_statistic_by_church_all` (
  `교회명` varchar(50) NOT NULL COMMENT '교회/지교회명',
  `교회형태` varchar(15) NOT NULL COMMENT '교회형태',
  `설립일` date DEFAULT NULL COMMENT '시작일',
  `직책` varchar(15) DEFAULT NULL COMMENT '직책',
  `관리자명` varchar(50) DEFAULT NULL COMMENT '한글이름',
  `전체1회` int(6) DEFAULT 0 COMMENT '전체1회',
  `전체4회` int(6) DEFAULT 0 COMMENT '전체4회',
  `학생1회` int(6) DEFAULT 0 COMMENT '학생이상1회',
  `학생4회` int(6) DEFAULT 0 COMMENT '학생이상4회',
  `학생반차` int(6) DEFAULT 0 COMMENT '학생이상 반차',
  `전체침례` int(6) DEFAULT 0 COMMENT '전체 침례',
  `전도인` int(6) DEFAULT 0 COMMENT '고정전도인',
  `지역장` int(6) DEFAULT 0 COMMENT '지역장',
  `구역장` int(6) DEFAULT 0 COMMENT '구역장',
  `관리부서` int(3) DEFAULT 0 COMMENT '관리부서',
  `연합회` varchar(15) DEFAULT NULL COMMENT '연합회',
  `본교회명` varchar(50) DEFAULT NULL COMMENT '본교회명',
  `정렬순서` int(10) NOT NULL DEFAULT 0 COMMENT '정렬순서',
  `교회개수` int(10) DEFAULT 0 COMMENT '교회개수'
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.temp_statistic_by_country 구조 내보내기
DROP TABLE IF EXISTS `temp_statistic_by_country`;
CREATE TABLE IF NOT EXISTS `temp_statistic_by_country` (
  `국가` varchar(30) NOT NULL,
  `소계` int(5) DEFAULT NULL,
  `MC` int(5) DEFAULT NULL,
  `BC` int(5) DEFAULT NULL,
  `PBC` int(5) DEFAULT NULL,
  `전체1회` int(10) DEFAULT NULL,
  `전체4회` int(10) DEFAULT NULL,
  `학생1회` int(10) DEFAULT NULL,
  `학생4회` int(10) DEFAULT NULL,
  `반차` int(7) DEFAULT NULL,
  `침례` int(5) DEFAULT NULL,
  `전도인` int(5) DEFAULT NULL,
  `지역장` int(5) DEFAULT NULL,
  `구역장` int(5) DEFAULT NULL,
  `인원` int(5) DEFAULT NULL,
  `당회장` int(5) DEFAULT NULL,
  `당회장대리` int(5) DEFAULT NULL,
  `동역` int(5) DEFAULT NULL,
  `생도` int(5) DEFAULT NULL,
  `지교회관리자` int(5) DEFAULT NULL,
  `예배소관리자` int(5) DEFAULT NULL,
  `당사모` int(5) DEFAULT NULL,
  `당대리사모` int(5) DEFAULT NULL,
  `동사모` int(5) DEFAULT NULL,
  `생도사모` int(5) DEFAULT NULL,
  `지관자사모` int(5) DEFAULT NULL,
  `예관자사모` int(5) DEFAULT NULL,
  `관리부서` int(3) NOT NULL DEFAULT 0 COMMENT '관리부서'
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 테이블 op_system.temp_statistic_by_pstaff 구조 내보내기
DROP TABLE IF EXISTS `temp_statistic_by_pstaff`;
CREATE TABLE IF NOT EXISTS `temp_statistic_by_pstaff` (
  `국가` varchar(15) NOT NULL,
  `연합회` varchar(15) DEFAULT NULL,
  `관리교회` varchar(50) NOT NULL,
  `관리자명` varchar(30) DEFAULT NULL,
  `전체1회` int(10) DEFAULT NULL,
  `소계` int(5) DEFAULT NULL,
  `당회장` int(5) DEFAULT NULL,
  `당회장대리` int(5) DEFAULT NULL,
  `동역` int(5) DEFAULT NULL,
  `예비생도` int(5) DEFAULT NULL,
  `지교회관리자` int(5) DEFAULT NULL,
  `예배소관리자` int(5) DEFAULT NULL,
  `침례권자` int(5) DEFAULT NULL,
  `목사` int(5) DEFAULT NULL,
  `장로` int(5) DEFAULT NULL,
  `전도사` int(5) DEFAULT NULL,
  `집사` int(5) DEFAULT NULL,
  `형제` int(5) DEFAULT NULL,
  `관리부서` int(3) NOT NULL DEFAULT 0 COMMENT '관리부서',
  `정렬순서` int(10) NOT NULL DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 내보낼 데이터가 선택되어 있지 않습니다.

-- 뷰 op_system.v0_history_church_temp 구조 내보내기
DROP VIEW IF EXISTS `v0_history_church_temp`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v0_history_church_temp` (
	`커스텀코드` INT(6) NOT NULL,
	`교회코드` VARCHAR(20) NOT NULL COLLATE 'utf8_general_ci',
	`교회명` VARCHAR(50) NULL COMMENT '교회명' COLLATE 'utf8_general_ci',
	`교회구분` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`관리교회코드` VARCHAR(20) NULL COMMENT '관리교회 코드' COLLATE 'utf8_general_ci',
	`관리교회명` VARCHAR(50) NULL COMMENT '교회명' COLLATE 'utf8_general_ci',
	`시작일` DATE NOT NULL,
	`종료일` DATE NULL,
	`선임일` DATE NULL,
	`생명번호` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`한글이름` VARCHAR(50) NOT NULL COLLATE 'utf8_general_ci',
	`직분` VARCHAR(50) NULL COLLATE 'utf8_general_ci'
) ENGINE=MyISAM;

-- 뷰 op_system.v0_pstaff_information 구조 내보내기
DROP VIEW IF EXISTS `v0_pstaff_information`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v0_pstaff_information` (
	`교회코드` VARCHAR(20) NULL COMMENT '교회코드' COLLATE 'utf8_general_ci',
	`교회명` VARCHAR(50) NULL COMMENT '교회명' COLLATE 'utf8_general_ci',
	`영문교회명` VARCHAR(80) NULL COMMENT '영문 교회명' COLLATE 'utf8_general_ci',
	`지교회명` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`영문지교회명` VARCHAR(80) NULL COLLATE 'utf8_general_ci',
	`선교국가` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`생명번호` VARCHAR(20) NOT NULL COMMENT '생명번호' COLLATE 'utf8_general_ci',
	`한글이름(직분)` VARCHAR(53) NOT NULL COLLATE 'utf8_general_ci',
	`영문이름` VARCHAR(50) NOT NULL COMMENT '영문이름' COLLATE 'utf8_general_ci',
	`직책` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`직책2` VARCHAR(15) NULL COMMENT '특수직책' COLLATE 'utf8_general_ci',
	`생년월일` DATE NOT NULL COMMENT '생년월일',
	`국적` VARCHAR(50) NOT NULL COMMENT '국적' COLLATE 'utf8_general_ci',
	`고향` VARCHAR(200) NULL COMMENT '본가' COLLATE 'utf8_general_ci',
	`사모고향` VARCHAR(200) NULL COMMENT '본가' COLLATE 'utf8_general_ci',
	`(해외)최초발령일` DATE NULL,
	`현당회발령일` DATE NULL,
	`배우자생번` VARCHAR(20) NULL COMMENT '생명번호' COLLATE 'utf8_general_ci',
	`사모한글이름(직분)` VARCHAR(53) NULL COLLATE 'utf8_general_ci',
	`사모영문이름` VARCHAR(50) NULL COMMENT '영문이름' COLLATE 'utf8_general_ci',
	`사모직책` VARCHAR(10) NULL COLLATE 'utf8_general_ci',
	`배우자 생년월일` DATE NULL COMMENT '생년월일',
	`생도기수` VARCHAR(10) NULL COLLATE 'utf8_general_ci',
	`관리부서` INT(3) NOT NULL COMMENT '관리부서',
	`사모국적` VARCHAR(20) NULL COMMENT '국적' COLLATE 'utf8_general_ci',
	`한글이름` VARCHAR(50) NOT NULL COMMENT '한글이름' COLLATE 'utf8_general_ci',
	`사모한글이름` VARCHAR(50) NULL COMMENT '한글이름' COLLATE 'utf8_general_ci',
	`교육` VARCHAR(100) NULL COMMENT '학력' COLLATE 'utf8_general_ci',
	`사모교육` VARCHAR(100) NULL COMMENT '학력' COLLATE 'utf8_general_ci',
	`직분` VARCHAR(50) NULL COMMENT '직분' COLLATE 'utf8_general_ci',
	`사모직분` VARCHAR(50) NULL COMMENT '직분' COLLATE 'utf8_general_ci',
	`지교회코드` VARCHAR(20) NULL COMMENT '교회코드' COLLATE 'utf8_general_ci'
) ENGINE=MyISAM;

-- 뷰 op_system.v0_pstaff_information_all 구조 내보내기
DROP VIEW IF EXISTS `v0_pstaff_information_all`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v0_pstaff_information_all` (
	`교회코드` VARCHAR(20) NULL COMMENT '교회코드' COLLATE 'utf8_general_ci',
	`교회명` VARCHAR(50) NULL COMMENT '교회명' COLLATE 'utf8_general_ci',
	`영문교회명` VARCHAR(80) NULL COMMENT '영문 교회명' COLLATE 'utf8_general_ci',
	`지교회명` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`영문지교회명` VARCHAR(80) NULL COLLATE 'utf8_general_ci',
	`선교국가` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`생명번호` VARCHAR(20) NOT NULL COMMENT '생명번호' COLLATE 'utf8_general_ci',
	`한글이름(직분)` VARCHAR(53) NOT NULL COLLATE 'utf8_general_ci',
	`영문이름` VARCHAR(50) NOT NULL COMMENT '영문이름' COLLATE 'utf8_general_ci',
	`직책` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`직책2` VARCHAR(15) NULL COMMENT '특수직책' COLLATE 'utf8_general_ci',
	`생년월일` DATE NOT NULL COMMENT '생년월일',
	`국적` VARCHAR(50) NOT NULL COMMENT '국적' COLLATE 'utf8_general_ci',
	`고향` VARCHAR(200) NULL COMMENT '본가' COLLATE 'utf8_general_ci',
	`사모고향` VARCHAR(200) NULL COMMENT '본가' COLLATE 'utf8_general_ci',
	`(해외)최초발령일` DATE NULL,
	`현당회발령일` DATE NULL,
	`배우자생번` VARCHAR(20) NULL COMMENT '생명번호' COLLATE 'utf8_general_ci',
	`사모한글이름(직분)` VARCHAR(53) NULL COLLATE 'utf8_general_ci',
	`사모영문이름` VARCHAR(50) NULL COMMENT '영문이름' COLLATE 'utf8_general_ci',
	`사모직책` VARCHAR(10) NULL COLLATE 'utf8_general_ci',
	`배우자 생년월일` DATE NULL COMMENT '생년월일',
	`생도기수` VARCHAR(10) NULL COLLATE 'utf8_general_ci',
	`관리부서` INT(3) NOT NULL COMMENT '관리부서',
	`사모국적` VARCHAR(20) NULL COMMENT '국적' COLLATE 'utf8_general_ci',
	`한글이름` VARCHAR(50) NOT NULL COMMENT '한글이름' COLLATE 'utf8_general_ci',
	`사모한글이름` VARCHAR(50) NULL COMMENT '한글이름' COLLATE 'utf8_general_ci',
	`교육` VARCHAR(100) NULL COMMENT '학력' COLLATE 'utf8_general_ci',
	`사모교육` VARCHAR(100) NULL COMMENT '학력' COLLATE 'utf8_general_ci',
	`직분` VARCHAR(50) NULL COMMENT '직분' COLLATE 'utf8_general_ci',
	`사모직분` VARCHAR(50) NULL COMMENT '직분' COLLATE 'utf8_general_ci',
	`지교회코드` VARCHAR(20) NULL COMMENT '교회코드' COLLATE 'utf8_general_ci'
) ENGINE=MyISAM;

-- 뷰 op_system.v0_theological_history 구조 내보내기
DROP VIEW IF EXISTS `v0_theological_history`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v0_theological_history` (
	`theological_cd` INT(5) NULL,
	`LifeNo` VARCHAR(20) NULL COMMENT '생명번호' COLLATE 'utf8_general_ci',
	`name_ko` VARCHAR(50) NOT NULL COMMENT '한글이름' COLLATE 'utf8_general_ci',
	`Level` VARCHAR(10) NULL COMMENT '예비생도 단계' COLLATE 'utf8_general_ci',
	`CUR_STATUS` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`Start_dt` DATE NULL COMMENT '시작일',
	`End_dt` DATE NULL COMMENT '종료일',
	`Resign_dt` DATE NULL COMMENT '성도복귀일',
	`church_sid` VARCHAR(15) NULL COMMENT '추천교회' COLLATE 'utf8_general_ci',
	`church_nm` VARCHAR(50) NULL COMMENT '교회명' COLLATE 'utf8_general_ci'
) ENGINE=MyISAM;

-- 뷰 op_system.v_atten_detail_churchlist 구조 내보내기
DROP VIEW IF EXISTS `v_atten_detail_churchlist`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v_atten_detail_churchlist` (
	`church_sid_custom` INT(6) NOT NULL,
	`start_dt` DATE NULL,
	`end_dt` DATE NULL,
	`church_sid` VARCHAR(20) NULL COLLATE 'utf8_general_ci',
	`church_nm` VARCHAR(50) NULL COMMENT '교회명' COLLATE 'utf8_general_ci',
	`church_gb` VARCHAR(15) NULL COMMENT '교회형태' COLLATE 'utf8_general_ci',
	`ovs_dept` VARCHAR(15) NULL COMMENT '해외국 관리부서' COLLATE 'utf8_general_ci'
) ENGINE=MyISAM;

-- 뷰 op_system.v_churchlist_final 구조 내보내기
DROP VIEW IF EXISTS `v_churchlist_final`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v_churchlist_final` (
	`교회커스텀코드` INT(6) NOT NULL,
	`교회코드` VARCHAR(20) NOT NULL COLLATE 'utf8_general_ci',
	`교회명(ko)` VARCHAR(50) NULL COMMENT '교회이름' COLLATE 'utf8_general_ci',
	`교회명(en)` VARCHAR(80) NULL COMMENT '영문 교회명' COLLATE 'utf8_general_ci',
	`교회구분` VARCHAR(15) NULL COMMENT '교회형태' COLLATE 'utf8_general_ci',
	`본교회코드` VARCHAR(20) NULL COMMENT '관리교회 코드' COLLATE 'utf8_general_ci',
	`본교회명` VARCHAR(50) NULL COMMENT '교회이름' COLLATE 'utf8_general_ci',
	`관리자` VARCHAR(100) NULL COLLATE 'utf8_general_ci',
	`관리자직분` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`관리자직책` VARCHAR(30) NULL COLLATE 'utf8_general_ci',
	`GEO코드` INT(10) NULL COMMENT 'GEO코드',
	`관리부서` VARCHAR(15) NULL COMMENT '해외국 관리부서' COLLATE 'utf8_general_ci',
	`위도` DECIMAL(15,10) NULL COMMENT '위도',
	`경도` DECIMAL(15,10) NULL COMMENT '경도',
	`논리삭제` TINYINT(1) NULL COMMENT '논리삭제',
	`정렬순서` INT(10) NULL COMMENT '정렬순서'
) ENGINE=MyISAM;

-- 뷰 op_system.v_churchlist_nomatch 구조 내보내기
DROP VIEW IF EXISTS `v_churchlist_nomatch`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v_churchlist_nomatch` (
	`church_sid` VARCHAR(20) NOT NULL COMMENT '교회코드' COLLATE 'utf8_general_ci',
	`church_nm` VARCHAR(50) NOT NULL COMMENT '교회명' COLLATE 'utf8_general_ci',
	`church_gb` VARCHAR(15) NOT NULL COMMENT '교회형태' COLLATE 'utf8_general_ci',
	`main_church` VARCHAR(50) NULL COMMENT '교회명' COLLATE 'utf8_general_ci',
	`start_dt` DATE NULL COMMENT '시작일',
	`end_dt` VARCHAR(10) NULL COLLATE 'utf8mb4_general_ci',
	`ovs_dept` VARCHAR(15) NULL COMMENT '해외국 관리부서' COLLATE 'utf8_general_ci'
) ENGINE=MyISAM;

-- 뷰 op_system.v_familyinfo 구조 내보내기
DROP VIEW IF EXISTS `v_familyinfo`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v_familyinfo` (
	`family_id` INT(5) NOT NULL COMMENT '구성원id',
	`family_cd` INT(5) NOT NULL COMMENT '가족코드',
	`relations` VARCHAR(9) NOT NULL COLLATE 'utf8_general_ci',
	`lifeno` VARCHAR(20) NULL COMMENT '생명번호' COLLATE 'utf8_general_ci',
	`name_ko` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`name_en` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`church_sid` VARCHAR(15) NULL COMMENT '교회' COLLATE 'utf8_general_ci',
	`church_nm` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`title` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`position` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`birthday` VARCHAR(10) NULL COLLATE 'utf8mb4_general_ci',
	`education` VARCHAR(100) NULL COLLATE 'utf8_general_ci',
	`religion` VARCHAR(20) NULL COLLATE 'utf8_general_ci',
	`recognition` VARCHAR(5) NULL COMMENT '본교인식' COLLATE 'utf8_general_ci',
	`memo` VARCHAR(300) NULL COMMENT '메모' COLLATE 'utf8_general_ci',
	`suspend` TINYINT(1) NOT NULL COMMENT '0-생존,1-별세',
	`churchFullName` VARCHAR(50) NULL COLLATE 'utf8_general_ci'
) ENGINE=MyISAM;

-- 뷰 op_system.v_history_church 구조 내보내기
DROP VIEW IF EXISTS `v_history_church`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v_history_church` (
	`커스텀코드` INT(11) NULL,
	`교회코드` VARCHAR(20) NULL COLLATE 'utf8_general_ci',
	`날짜` DATE NULL,
	`생명번호` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`교회연혁` VARCHAR(200) NULL COLLATE 'utf8_general_ci'
) ENGINE=MyISAM;

-- 뷰 op_system.v_phone 구조 내보내기
DROP VIEW IF EXISTS `v_phone`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v_phone` (
	`선교국가` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`시차` TIME NULL COMMENT '시차',
	`교회코드` INT(6) NULL,
	`교회명` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`유선전화` VARCHAR(50) NULL COMMENT '유선전화' COLLATE 'utf8_general_ci',
	`인터넷전화` VARCHAR(50) NULL COMMENT '인터넷전화' COLLATE 'utf8_general_ci',
	`선지자전화번호` VARCHAR(80) NULL COMMENT '전화번호' COLLATE 'utf8_general_ci',
	`배우자전화번호` VARCHAR(80) NULL COMMENT '전화번호' COLLATE 'utf8_general_ci',
	`선지자생명번호` VARCHAR(20) NOT NULL COMMENT '생명번호' COLLATE 'utf8_general_ci',
	`한글이름(직분)` VARCHAR(53) NOT NULL COLLATE 'utf8_general_ci',
	`직책` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`배우자생명번호` VARCHAR(20) NULL COMMENT '생명번호' COLLATE 'utf8_general_ci',
	`사모한글이름(직분)` VARCHAR(53) NULL COLLATE 'utf8_general_ci',
	`사모직책` VARCHAR(10) NULL COLLATE 'utf8_general_ci',
	`관리교회명` VARCHAR(50) NULL COMMENT '교회명' COLLATE 'utf8_general_ci',
	`관리부서` VARCHAR(15) NULL COMMENT '해외국 관리부서' COLLATE 'utf8_general_ci',
	`영문이름` VARCHAR(53) NOT NULL COLLATE 'utf8_general_ci',
	`사모영문이름` VARCHAR(53) NULL COLLATE 'utf8_general_ci',
	`교회주소` VARCHAR(1000) NULL COLLATE 'utf8_general_ci',
	`본교회코드` VARCHAR(20) NULL COMMENT '교회코드' COLLATE 'utf8_general_ci',
	`지교회코드` VARCHAR(20) NULL COLLATE 'utf8_general_ci',
	`정렬순서` INT(10) NULL COMMENT '정렬순서',
	`생년월일` DATE NOT NULL COMMENT '생년월일',
	`최초발령일` DATE NULL,
	`영문교회명` VARCHAR(100) NULL COLLATE 'utf8_general_ci'
) ENGINE=MyISAM;

-- 뷰 op_system.v_phone_export 구조 내보내기
DROP VIEW IF EXISTS `v_phone_export`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v_phone_export` (
	`선교국가` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`시차` VARCHAR(10) NULL COLLATE 'utf8mb4_general_ci',
	`교회명` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`인터넷전화` VARCHAR(50) NULL COMMENT '인터넷전화' COLLATE 'utf8_general_ci',
	`유선전화` VARCHAR(50) NULL COMMENT '유선전화' COLLATE 'utf8_general_ci',
	`직책` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`한글이름(직분)` VARCHAR(53) NOT NULL COLLATE 'utf8_general_ci',
	`선지자전화번호` VARCHAR(80) NULL COMMENT '전화번호' COLLATE 'utf8_general_ci',
	`사모한글이름(직분)` VARCHAR(53) NULL COLLATE 'utf8_general_ci',
	`배우자전화번호` VARCHAR(80) NULL COMMENT '전화번호' COLLATE 'utf8_general_ci',
	`관리부서` VARCHAR(15) NULL COMMENT '해외국 관리부서' COLLATE 'utf8_general_ci'
) ENGINE=MyISAM;

-- 뷰 op_system.v_pstaff_detail 구조 내보내기
DROP VIEW IF EXISTS `v_pstaff_detail`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v_pstaff_detail` (
	`본교회명` VARCHAR(50) NULL COMMENT '교회명' COLLATE 'utf8_general_ci',
	`지교회명` VARCHAR(50) NULL COMMENT '교회명' COLLATE 'utf8_general_ci',
	`한글이름` VARCHAR(50) NOT NULL COMMENT '한글이름' COLLATE 'utf8_general_ci',
	`영문이름` VARCHAR(50) NOT NULL COMMENT '영문이름' COLLATE 'utf8_general_ci',
	`직분` VARCHAR(50) NULL COMMENT '직분' COLLATE 'utf8_general_ci',
	`직책` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`제2직책` VARCHAR(15) NULL COMMENT '특수직책' COLLATE 'utf8_general_ci',
	`생년월일` DATE NOT NULL COMMENT '생년월일',
	`생명번호` VARCHAR(20) NOT NULL COMMENT '생명번호' COLLATE 'utf8_general_ci',
	`학력` VARCHAR(100) NULL COMMENT '학력' COLLATE 'utf8_general_ci',
	`침례권` VARCHAR(2) NOT NULL COMMENT '침례권' COLLATE 'utf8_general_ci',
	`안수일` DATE NULL COMMENT '침례권안수',
	`유급` VARCHAR(3) NULL COLLATE 'utf8mb4_general_ci',
	`건강` TEXT(65535) NULL COMMENT '건강사항' COLLATE 'utf8_general_ci',
	`최초출국일` DATE NULL COMMENT '해외 최초 발령일',
	`국적` VARCHAR(50) NOT NULL COMMENT '국적' COLLATE 'utf8_general_ci',
	`비자` VARCHAR(50) NULL COMMENT '비자종류' COLLATE 'utf8_general_ci',
	`비자만료일` DATE NULL COMMENT '종료일',
	`본가위치` VARCHAR(200) NULL COMMENT '본가' COLLATE 'utf8_general_ci',
	`가족사항` VARCHAR(700) NULL COMMENT '가족사항' COLLATE 'utf8_general_ci',
	`사모한글이름` VARCHAR(50) NULL COMMENT '한글이름' COLLATE 'utf8_general_ci',
	`사모영문이름` VARCHAR(50) NULL COMMENT '영문이름' COLLATE 'utf8_general_ci',
	`사모직분` VARCHAR(50) NULL COMMENT '직분' COLLATE 'utf8_general_ci',
	`사모직책` VARCHAR(10) NULL COMMENT '사모직책' COLLATE 'utf8_general_ci',
	`사모생년월일` DATE NULL COMMENT '생년월일',
	`사모생명번호` VARCHAR(20) NULL COMMENT '생명번호' COLLATE 'utf8_general_ci',
	`사모학력` VARCHAR(100) NULL COMMENT '학력' COLLATE 'utf8_general_ci',
	`자녀1생명번호` VARCHAR(20) NULL COMMENT '자녀1 생명번호' COLLATE 'utf8_general_ci',
	`자녀1생년월일` DATE NULL COMMENT '자녀1 생년월일',
	`자녀2생명번호` VARCHAR(20) NULL COMMENT '자녀2 생명번호' COLLATE 'utf8_general_ci',
	`자녀2생년월일` DATE NULL COMMENT '자녀2 생년월일',
	`자녀3생명번호` VARCHAR(20) NULL COMMENT '자녀2 생명번호' COLLATE 'utf8_general_ci',
	`자녀3생년월일` DATE NULL COMMENT '자녀2 생년월일',
	`사모국적` VARCHAR(20) NULL COMMENT '국적' COLLATE 'utf8_general_ci',
	`사모건강` TEXT(65535) NULL COMMENT '건강사항' COLLATE 'utf8_general_ci',
	`사모비자` VARCHAR(50) NULL COMMENT '비자종류' COLLATE 'utf8_general_ci',
	`사모비자만료일` DATE NULL COMMENT '종료일',
	`친정위치` VARCHAR(200) NULL COMMENT '본가' COLLATE 'utf8_general_ci',
	`사모가족사항` VARCHAR(700) NULL COMMENT '가족사항' COLLATE 'utf8_general_ci',
	`발표주제` INT(3) NULL COMMENT '발표개수',
	`평균점수` DECIMAL(5,2) NULL COMMENT '발표점수',
	`생도기수` INT(4) NULL COMMENT '한국인 생도기수',
	`예비생도단계` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`선교국가` VARCHAR(50) NULL COMMENT '국가명(한글)' COLLATE 'utf8_general_ci',
	`현당회발령일` DATE NULL
) ENGINE=MyISAM;

-- 뷰 op_system.v_pstaff_detail_accomplishment 구조 내보내기
DROP VIEW IF EXISTS `v_pstaff_detail_accomplishment`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v_pstaff_detail_accomplishment` (
	`교회코드` VARCHAR(15) NOT NULL COLLATE 'utf8_general_ci',
	`교회명` VARCHAR(53) NULL COLLATE 'utf8_general_ci',
	`생명번호` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`한글이름` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`날짜` DATE NULL,
	`전체1회` INT(11) NULL,
	`전체4회` INT(11) NULL,
	`학생1회` INT(11) NULL,
	`학생4회` INT(11) NULL,
	`반차` INT(11) NULL,
	`침례` INT(11) NULL,
	`전도인` INT(11) NULL,
	`구역장` INT(11) NULL,
	`지역장` INT(11) NULL,
	`직분` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`직책` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`관리시작일` DATE NOT NULL,
	`관리종료일` DATE NOT NULL,
	`교회구분` VARCHAR(15) NULL COLLATE 'utf8_general_ci'
) ENGINE=MyISAM;

-- 뷰 op_system.v_pstaff_detail_accomplishment_both 구조 내보내기
DROP VIEW IF EXISTS `v_pstaff_detail_accomplishment_both`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v_pstaff_detail_accomplishment_both` (
	`교회코드` VARCHAR(15) NOT NULL COLLATE 'utf8_general_ci',
	`교회명` VARCHAR(53) NULL COLLATE 'utf8_general_ci',
	`생명번호` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`한글이름` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`날짜` DATE NULL,
	`전체1회` INT(11) NULL,
	`전체4회` INT(11) NULL,
	`학생1회` INT(11) NULL,
	`학생4회` INT(11) NULL,
	`반차` INT(11) NULL,
	`침례` INT(11) NULL,
	`전도인` INT(11) NULL,
	`구역장` INT(11) NULL,
	`지역장` INT(11) NULL,
	`직분` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`직책` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`관리시작일` DATE NOT NULL,
	`관리종료일` DATE NOT NULL,
	`교회구분` VARCHAR(15) NULL COLLATE 'utf8_general_ci'
) ENGINE=MyISAM;

-- 뷰 op_system.v_pstaff_detail_accomplishment_main 구조 내보내기
DROP VIEW IF EXISTS `v_pstaff_detail_accomplishment_main`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v_pstaff_detail_accomplishment_main` (
	`교회코드` VARCHAR(15) NOT NULL COLLATE 'utf8_general_ci',
	`교회명` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`생명번호` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`한글이름` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`날짜` DATE NULL,
	`전체1회` INT(11) NULL,
	`전체4회` INT(11) NULL,
	`학생1회` INT(11) NULL,
	`학생4회` INT(11) NULL,
	`반차` INT(11) NULL,
	`침례` INT(11) NULL,
	`전도인` INT(11) NULL,
	`구역장` INT(11) NULL,
	`지역장` INT(11) NULL,
	`직분` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`직책` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`관리시작일` DATE NOT NULL,
	`관리종료일` DATE NOT NULL,
	`교회구분` VARCHAR(15) NULL COLLATE 'utf8_general_ci'
) ENGINE=MyISAM;

-- 뷰 op_system.v_pstaff_detail_concise_transfer_history 구조 내보내기
DROP VIEW IF EXISTS `v_pstaff_detail_concise_transfer_history`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v_pstaff_detail_concise_transfer_history` (
	`교회명` VARCHAR(53) NULL COLLATE 'utf8_general_ci',
	`시작일` DATE NULL,
	`종료일` DATE NULL,
	`관리시작일` DATE NOT NULL,
	`관리종료일` DATE NOT NULL,
	`기간` INT(6) NULL,
	`직분` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`직책` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`교회구분` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`생명번호` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`교회코드` VARCHAR(15) NOT NULL COLLATE 'utf8_general_ci'
) ENGINE=MyISAM;

-- 뷰 op_system.v_pstaff_detail_concise_transfer_history_both 구조 내보내기
DROP VIEW IF EXISTS `v_pstaff_detail_concise_transfer_history_both`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v_pstaff_detail_concise_transfer_history_both` (
	`교회명` VARCHAR(53) NULL COLLATE 'utf8_general_ci',
	`시작일` DATE NULL,
	`종료일` DATE NULL,
	`관리시작일` DATE NOT NULL,
	`관리종료일` DATE NOT NULL,
	`기간` INT(6) NULL,
	`직분` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`직책` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`교회구분` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`생명번호` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`교회코드` VARCHAR(15) NOT NULL COLLATE 'utf8_general_ci'
) ENGINE=MyISAM;

-- 뷰 op_system.v_pstaff_detail_concise_transfer_history_main 구조 내보내기
DROP VIEW IF EXISTS `v_pstaff_detail_concise_transfer_history_main`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v_pstaff_detail_concise_transfer_history_main` (
	`교회명` VARCHAR(53) NULL COLLATE 'utf8_general_ci',
	`시작일` DATE NULL,
	`종료일` DATE NULL,
	`관리시작일` DATE NOT NULL,
	`관리종료일` DATE NOT NULL,
	`기간` INT(6) NULL,
	`직분` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`직책` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`교회구분` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`생명번호` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`교회코드` VARCHAR(15) NOT NULL COLLATE 'utf8_general_ci'
) ENGINE=MyISAM;

-- 뷰 op_system.v_pstaff_detail_flight 구조 내보내기
DROP VIEW IF EXISTS `v_pstaff_detail_flight`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v_pstaff_detail_flight` (
	`생명번호` VARCHAR(20) NOT NULL COLLATE 'utf8_general_ci',
	`방문일자` DATE NOT NULL,
	`방문목적` VARCHAR(100) NULL COLLATE 'utf8_general_ci'
) ENGINE=MyISAM;

-- 뷰 op_system.v_pstaff_detail_title 구조 내보내기
DROP VIEW IF EXISTS `v_pstaff_detail_title`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v_pstaff_detail_title` (
	`생명번호` VARCHAR(20) NOT NULL COMMENT '생명번호' COLLATE 'utf8_general_ci',
	`교회명` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`임명일` DATE NOT NULL COMMENT '시작일',
	`직분` VARCHAR(50) NOT NULL COMMENT '직분' COLLATE 'utf8_general_ci'
) ENGINE=MyISAM;

-- 뷰 op_system.v_pstaff_detail_transfer 구조 내보내기
DROP VIEW IF EXISTS `v_pstaff_detail_transfer`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v_pstaff_detail_transfer` (
	`생명번호` VARCHAR(50) NOT NULL COLLATE 'utf8_general_ci',
	`발령일` DATE NOT NULL,
	`직분/직책` VARCHAR(66) NULL COLLATE 'utf8_general_ci',
	`교회구분` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`교회명` VARCHAR(153) NULL COLLATE 'utf8_general_ci',
	`교회코드` VARCHAR(20) NULL COLLATE 'utf8_general_ci',
	`기간` VARCHAR(96) NULL COLLATE 'utf8mb4_general_ci',
	`이력구분` VARCHAR(8) NOT NULL COLLATE 'utf8mb4_general_ci',
	`종료일` DATE NULL,
	`직분` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`직책` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`국가명` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`현재국가` VARCHAR(50) NULL COMMENT '국가명(한글)' COLLATE 'utf8_general_ci'
) ENGINE=MyISAM;

-- 뷰 op_system.v_search_titleposition 구조 내보내기
DROP VIEW IF EXISTS `v_search_titleposition`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v_search_titleposition` (
	`생명번호` VARCHAR(20) NOT NULL COMMENT '생명번호' COLLATE 'utf8_general_ci',
	`교회명` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`영문교회명` VARCHAR(80) NULL COMMENT '영문 교회명' COLLATE 'utf8_general_ci',
	`지교회명` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`영문지교회명` VARCHAR(80) NULL COLLATE 'utf8_general_ci',
	`선교국가` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`한글이름(직분)` VARCHAR(53) NOT NULL COLLATE 'utf8_general_ci',
	`영문이름` VARCHAR(50) NOT NULL COMMENT '영문이름' COLLATE 'utf8_general_ci',
	`직책` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`직책2` VARCHAR(15) NULL COMMENT '특수직책' COLLATE 'utf8_general_ci',
	`생년월일` DATE NOT NULL COMMENT '생년월일',
	`국적` VARCHAR(50) NOT NULL COMMENT '국적' COLLATE 'utf8_general_ci',
	`(해외)최초발령일` VARCHAR(10) NOT NULL COLLATE 'utf8mb4_general_ci',
	`현당회발령일` VARCHAR(10) NOT NULL COLLATE 'utf8mb4_general_ci',
	`배우자생번` VARCHAR(20) NULL COLLATE 'utf8_general_ci',
	`사모한글이름(직분)` VARCHAR(53) NULL COLLATE 'utf8_general_ci',
	`사모영문이름` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`사모직책` VARCHAR(10) NULL COLLATE 'utf8_general_ci',
	`사모직책2` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`배우자 생년월일` DATE NULL,
	`생도기수` VARCHAR(10) NULL COLLATE 'utf8_general_ci',
	`직분` VARCHAR(50) NULL COMMENT '직분' COLLATE 'utf8_general_ci',
	`사모직분` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`침례권` VARCHAR(2) NOT NULL COMMENT '침례권' COLLATE 'utf8_general_ci',
	`연합회` VARCHAR(10) NULL COMMENT '연합회명' COLLATE 'utf8_general_ci',
	`전체1회` INT(6) NULL COMMENT '전체1회',
	`학생1회` INT(6) NULL COMMENT '학생이상1회',
	`지교회전체1회` INT(11) NULL,
	`지교회학생1회` INT(11) NULL,
	`관리지교회` BIGINT(21) NULL,
	`관리예배소` BIGINT(21) NULL,
	`동역` BIGINT(21) NULL,
	`지교회관리자` BIGINT(21) NULL,
	`예배소관리자` BIGINT(21) NULL,
	`예비생도` BIGINT(21) NULL,
	`사모국적` VARCHAR(20) NULL COMMENT '국적' COLLATE 'utf8_general_ci',
	`직책2시작일` DATE NULL COMMENT '시작일',
	`사모직책2시작일` DATE NULL COMMENT '시작일',
	`전체1회(2달 전)` INT(6) NULL COMMENT '전체1회',
	`학생1회(2달 전)` INT(6) NULL COMMENT '학생이상1회',
	`지교회전체1회(2달 전)` INT(11) NULL,
	`지교회학생1회(2달 전)` INT(11) NULL,
	`관리부서` INT(3) NOT NULL COMMENT '관리부서',
	`연합회 정렬순서` INT(3) NULL COMMENT '정렬순서',
	`본교회 정렬순서` INT(10) NULL COMMENT '정렬순서',
	`교회구분` VARCHAR(22) NULL COLLATE 'utf8_general_ci',
	`지교회 정렬순서` INT(10) NULL COMMENT '정렬순서',
	`교회명(전체)` VARCHAR(50) NULL COMMENT '교회명' COLLATE 'utf8_general_ci',
	`지교회역할` VARCHAR(30) NULL COLLATE 'utf8_general_ci'
) ENGINE=MyISAM;

-- 뷰 op_system.v_transfer_history 구조 내보내기
DROP VIEW IF EXISTS `v_transfer_history`;
-- VIEW 종속성 오류를 극복하기 위해 임시 테이블을 생성합니다.
CREATE TABLE `v_transfer_history` (
	`생명번호` VARCHAR(20) NOT NULL COLLATE 'utf8_general_ci',
	`교회명` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`선지자이름(직분)` VARCHAR(53) NULL COLLATE 'utf8_general_ci',
	`선지자직책` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`선지자생년월일` DATE NULL,
	`선지자국적` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`선지자고향` VARCHAR(200) NULL COLLATE 'utf8_general_ci',
	`선지자가족` VARCHAR(700) NULL COLLATE 'utf8_general_ci',
	`선지자건강` MEDIUMTEXT NULL COLLATE 'utf8_general_ci',
	`선지자기타` MEDIUMTEXT NULL COLLATE 'utf8_general_ci',
	`선지자마지막방문일` DATE NULL,
	`선지자방문목적` VARCHAR(100) NULL COLLATE 'utf8_general_ci',
	`사모생번` VARCHAR(20) NULL COLLATE 'utf8_general_ci',
	`사모이름(직분)` VARCHAR(53) NULL COLLATE 'utf8_general_ci',
	`사모직책` VARCHAR(10) NULL COLLATE 'utf8_general_ci',
	`사모생년월일` DATE NULL,
	`사모국적` VARCHAR(20) NULL COLLATE 'utf8_general_ci',
	`사모고향` VARCHAR(200) NULL COLLATE 'utf8_general_ci',
	`사모가족` VARCHAR(700) NULL COLLATE 'utf8_general_ci',
	`사모건강` MEDIUMTEXT NULL COLLATE 'utf8_general_ci',
	`사모기타` MEDIUMTEXT NULL COLLATE 'utf8_general_ci',
	`사모마지막방문일` DATE NULL,
	`사모방문목적` VARCHAR(100) NULL COLLATE 'utf8_general_ci',
	`혼인일` DATE NULL,
	`(해외)최초발령일` DATE NULL,
	`현당회발령일` DATE NULL,
	`발령일` DATE NOT NULL,
	`전출직책` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`전출직분` VARCHAR(1) NULL COLLATE 'utf8_general_ci',
	`전출교회` VARCHAR(53) NULL COLLATE 'utf8_general_ci',
	`전입직책` VARCHAR(15) NULL COLLATE 'utf8_general_ci',
	`전입직분` VARCHAR(1) NULL COLLATE 'utf8_general_ci',
	`전입교회` VARCHAR(50) NULL COLLATE 'utf8_general_ci',
	`자녀1 생명번호` VARCHAR(20) NULL COLLATE 'utf8_general_ci',
	`자녀1 생년월일` DATE NULL,
	`자녀2 생명번호` VARCHAR(20) NULL COLLATE 'utf8_general_ci',
	`자녀2 생년월일` DATE NULL,
	`자녀3 생명번호` VARCHAR(20) NULL COLLATE 'utf8_general_ci',
	`자녀3 생년월일` DATE NULL
) ENGINE=MyISAM;

-- 뷰 op_system.report_division_by_country 구조 내보내기
DROP VIEW IF EXISTS `report_division_by_country`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `report_division_by_country`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `report_division_by_country` AS SELECT
	DISTINCT 
	dept.dept_nm
	,users.user_nm
	,geo.country_nm_ko
FROM op_system.db_geodata geo
INNER JOIN op_system.db_ovs_dept dept
	ON geo.division = dept.dept_nm
INNER JOIN common.users users
	ON users.user_dept = dept.dept_id
		AND users.user_nm <> '이승언'
INNER JOIN op_system.a_auth_table auth
	ON auth.user_id = users.user_id
INNER JOIN op_system.a_authority auth_name
	ON auth.authority_id = auth_name.id
		AND auth_name.authority = 'SECTION_CHIEF'
ORDER BY dept.dept_nm, geo.country_nm_ko ;

-- 뷰 op_system.v0_history_church_temp 구조 내보내기
DROP VIEW IF EXISTS `v0_history_church_temp`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v0_history_church_temp`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v0_history_church_temp` AS SELECT 
  esta.church_sid_custom AS '커스텀코드', 
  esta.church_sid AS '교회코드', 
  churchlist.church_nm AS '교회명', 
  REPLACE(
    churchlist.church_gb, 'HBC', 
    'MC'
  ) AS '교회구분', 
  churchlist.main_church_cd AS '관리교회코드', 
  mainchurch.church_nm AS '관리교회명', 
  esta.start_dt AS '시작일', 
  esta.end_dt AS '종료일', 
  Ifnull(
    bcleader.start_dt, transfer.start_dt
  ) AS '선임일', 
  Ifnull(
    bcleader.lifeno, transfer.lifeno
  ) AS '생명번호', 
  IF(
    Ifnull(
      bcleader.lifeno, transfer.lifeno
    ) IS NOT NULL, 
    Ifnull(
      pstaff.name_ko, '일반식구'
    ), 
    '관리자 없음'
  ) AS '한글이름', 
  Ifnull(
    title.title, 
    IF(
      Ifnull(
        bcleader.lifeno, transfer.lifeno
      ) IS NULL, 
      NULL, 
      IF(
        Substr(
          Ifnull(
            bcleader.lifeno, transfer.lifeno
          ), 
          12, 
          1
        ) = '1', 
        '형제', 
        '자매'
      )
    )
  ) AS '직분' 
FROM 
  op_system.db_history_church_establish esta 
  LEFT JOIN op_system.db_churchlist churchlist ON(
    esta.church_sid = churchlist.church_sid
  ) 
  LEFT JOIN op_system.db_churchlist mainchurch ON(
    churchlist.main_church_cd = mainchurch.church_sid
  ) 
  LEFT JOIN op_system.db_branchleader bcleader ON(
    esta.church_sid = bcleader.church_sid 
    AND bcleader.responsibility NOT LIKE '%단순소속%'
  ) 
  LEFT JOIN (
    SELECT 
      a.church_sid AS church_sid, 
      b.lifeno AS lifeno, 
      a.start_dt AS start_dt 
    FROM 
      (
        (
          op_system.db_transfer a 
          LEFT JOIN op_system.db_pastoralstaff b ON(a.lifeno = b.lifeno)
        ) 
        LEFT JOIN op_system.db_position c ON(
          b.lifeno = c.lifeno 
          AND Last_day(a.start_dt) + INTERVAL 1 day BETWEEN c.start_dt 
          AND c.end_dt
        )
      ) 
    WHERE 
      c.position LIKE '%당%' 
    UNION 
    SELECT 
      a.presentChurch, 
      a.LifeNo, 
      a.Start_dt 
    FROM 
      (
        SELECT 
          pos.*, 
          pstaff.ovs_dept, 
          pstaff.name_ko AS presentOverseer, 
          LAG(pstaff.name_ko, 1) over (
            PARTITION BY pos.LifeNo 
            ORDER BY 
              pos.Start_dt
          ) AS lastOverseer, 
          LAG(pos.position, 1) over (
            PARTITION BY pos.LifeNo 
            ORDER BY 
              pos.Start_dt
          ) AS lastPosition, 
          presentChurch.church_sid AS presentChurch, 
          lastChurch.church_sid AS lastChurch 
        FROM 
          op_system.db_position pos 
          LEFT JOIN op_system.db_pastoralstaff pstaff ON pstaff.lifeno = pos.LifeNo 
          LEFT JOIN op_system.db_transfer presentChurch ON presentChurch.lifeno = pstaff.lifeno 
          AND pos.Start_dt BETWEEN presentChurch.start_dt 
          AND presentChurch.end_dt 
          LEFT JOIN op_system.db_transfer lastChurch ON lastChurch.lifeno = pstaff.lifeno 
          AND ADDDATE(pos.Start_dt, INTERVAL -7 DAY) BETWEEN lastChurch.start_dt 
          AND lastChurch.end_dt
      ) AS a 
    WHERE 
      a.Position LIKE '%당%' 
      AND a.presentChurch = a.lastChurch 
      AND a.presentOverseer = a.lastOverseer 
      AND a.lastPosition NOT LIKE '%당%'
  ) transfer ON(
    transfer.church_sid = esta.church_sid
  ) 
  LEFT JOIN op_system.db_pastoralstaff pstaff ON(
    Ifnull(
      bcleader.lifeno, transfer.lifeno
    ) = pstaff.lifeno
  ) 
  LEFT JOIN op_system.db_title title ON(
    pstaff.lifeno = title.lifeno 
    AND Ifnull(
      bcleader.start_dt, transfer.start_dt
    ) BETWEEN title.start_dt 
    AND title.end_dt
  ) 
WHERE 
  esta.start_dt >= '1995-01-01' 
ORDER BY 
  esta.church_sid_custom, 
  esta.start_dt, 
  Ifnull(
    bcleader.start_dt, transfer.start_dt
  ) ;

-- 뷰 op_system.v0_pstaff_information 구조 내보내기
DROP VIEW IF EXISTS `v0_pstaff_information`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v0_pstaff_information`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v0_pstaff_information` AS SELECT 
  churchlist.church_sid AS '교회코드', 
  churchlist.church_nm AS '교회명', 
  church_admin.church_nm_en AS '영문교회명', 
  Ifnull(
    branchlist.church_nm, churchlist.church_nm
  ) AS '지교회명', 
  Ifnull(
    branch_admin.church_nm_en, church_admin.church_nm_en
  ) AS '영문지교회명', 
  --  church_admin.country AS 선교국가, 
  IFNULL(
    geoBranch.country_nm_ko, geo.country_nm_ko
  ) AS '선교국가', 
  pstaff.lifeno AS '생명번호', 
  Concat(
    pstaff.name_ko, 
    Ifnull(
      Concat(
        '(', 
        LEFT(REPLACE(title.title,'봉사',''), 1), 
        ')'
      ), 
      ''
    )
  ) AS '한글이름(직분)', 
  pstaff.name_en AS '영문이름', 
  Ifnull(
    IF(
      position.position LIKE '%관리자%' 
      OR position.position LIKE '%당%' 
      OR position.position LIKE '%동%', 
      position.position, 
      theological.level
    ), 
    position2.position2
  ) AS '직책', 
  position2.position2 AS '직책2', 
  pstaff.birthday AS '생년월일', 
  pstaff.nationality AS '국적', 
  pstaff.home AS '고향', 
  spouse.home AS '사모고향', 
  IF(
    pstaff.appo_ovs IS NULL, appoint.start_dt, 
    pstaff.appo_ovs
  ) AS '(해외)최초발령일', 
--  belong.start_dt AS '현당회발령일', 
  IF(`position`.`position` IN ( '당회장', '당회장대리','동역' ),
     CASE 
		  WHEN `position`.`position` = '동역'
        THEN IF(`position`.`start_dt` >= `belong`.`start_dt`, `position`.`start_dt`, `belong`.`start_dt`)
        ELSE `belong`.`start_dt`
     END
     , NULL) AS '현당회발령일',
  spouse.lifeno AS '배우자생번', 
  Concat(
    spouse.name_ko, 
    Ifnull(
      Concat(
        '(', 
        LEFT(title_spouse.title, 1), 
        ')'
      ), 
      ''
    )
  ) AS '사모한글이름(직분)', 
  spouse.name_en AS '사모영문이름', 
  IF(
    spouse.lifeno IS NOT NULL, spouseposition.position_spouse, 
    NULL
  ) AS '사모직책', 
  spouse.birthday AS '배우자 생년월일', 
  IF(
    Ifnull(
      position.position, theological.level
    ) LIKE '당%' 
    OR Ifnull(
      position.position, theological.level
    ) LIKE '동%', 
    NULL, 
    theological.level
  ) AS '생도기수', 
  pstaff.ovs_dept AS '관리부서',
  spouse.nationality AS '사모국적', 
  pstaff.name_ko AS '한글이름', 
  spouse.name_ko AS '사모한글이름', 
  pstaff.education AS '교육', 
  spouse.education AS '사모교육', 
  title.title AS '직분', 
  title_spouse.title AS '사모직분', 
  branchlist.church_sid AS '지교회코드' 
FROM 
  op_system.db_pastoralstaff pstaff 
  LEFT JOIN op_system.db_title title ON(
    pstaff.lifeno = title.lifeno 
    AND Curdate() BETWEEN title.start_dt 
    AND title.end_dt
  ) 
  LEFT JOIN op_system.db_position position ON(
    pstaff.lifeno = position.lifeno 
    AND Curdate() BETWEEN position.start_dt 
    AND position.end_dt
  ) 
  LEFT JOIN op_system.db_position2 position2 ON(
    pstaff.lifeno = position2.lifeno 
    AND Curdate() BETWEEN position2.start_dt 
    AND position2.end_dt
  ) 
  LEFT JOIN op_system.db_transfer belong ON(
    pstaff.lifeno = belong.lifeno 
    AND Curdate() BETWEEN belong.start_dt 
    AND belong.end_dt
  ) 
  LEFT JOIN op_system.db_churchlist churchlist ON(
    belong.church_sid = churchlist.church_sid
  ) 
  LEFT JOIN op_system.db_branchleader belongbranch ON(
    pstaff.lifeno = belongbranch.lifeno 
    AND Curdate() BETWEEN belongbranch.start_dt 
    AND belongbranch.end_dt
  ) 
  LEFT JOIN (
    SELECT 
      * 
    FROM 
      op_system.db_churchlist
  ) branchlist ON(
    belongbranch.church_sid = branchlist.church_sid
  ) 
  LEFT JOIN op_system.db_theological theological ON(
    pstaff.lifeno = theological.lifeno 
    AND Curdate() BETWEEN theological.start_dt 
    AND theological.end_dt
  ) 
  LEFT JOIN op_system.db_pastoralwife spouse ON(
    pstaff.lifeno = spouse.lifeno_spouse
  ) 
  LEFT JOIN op_system.a_position_spouse spouseposition ON(
    IF(
      position.position LIKE '%관리자%' 
      OR position.position LIKE '%당%' 
      OR position.position LIKE '%동%', 
      position.position, 
      theological.level
    ) = spouseposition.position
  ) 
  LEFT JOIN (
    SELECT 
      op_system.db_position.lifeno AS lifeno, 
      Min(op_system.db_position.start_dt) AS start_dt, 
      op_system.db_position.position AS Position 
    FROM 
      op_system.db_position 
    WHERE 
      op_system.db_position.position = '당회장' 
      OR op_system.db_position.position = '당회장대리' 
      OR op_system.db_position.position = '동역' 
    GROUP BY 
      op_system.db_position.lifeno
  ) appoint ON(appoint.lifeno = pstaff.lifeno) 
  LEFT JOIN op_system.db_title title_spouse ON(
    title_spouse.lifeno = spouse.lifeno 
    AND Curdate() BETWEEN title_spouse.start_dt 
    AND title_spouse.end_dt
  ) 
  LEFT JOIN op_system.a_churchlist_admin church_admin ON(
    churchlist.church_sid = church_admin.church_sid
  ) 
  LEFT JOIN op_system.a_churchlist_admin branch_admin ON(
    branchlist.church_sid = branch_admin.church_sid
  ) 
  LEFT JOIN op_system.db_geodata geo ON geo.geo_cd = churchlist.geo_cd 
  LEFT JOIN op_system.db_geodata geoBranch ON geoBranch.geo_cd = branchlist.geo_cd 
WHERE 
  Ifnull(
    IF(
      position.position LIKE '%관리자%' 
      OR position.position LIKE '%당%' 
      OR position.position LIKE '%동%', 
      position.position, 
      theological.level
    ), 
    position2.position2
  ) IS NOT NULL 
GROUP BY 
  pstaff.lifeno 
ORDER BY 
  churchlist.church_nm, 
  IF(
    pstaff.appo_ovs IS NULL, appoint.start_dt, 
    pstaff.appo_ovs
  ) ;

-- 뷰 op_system.v0_pstaff_information_all 구조 내보내기
DROP VIEW IF EXISTS `v0_pstaff_information_all`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v0_pstaff_information_all`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v0_pstaff_information_all` AS SELECT 
  churchlist.church_sid AS '교회코드',
  churchlist.church_nm AS '교회명', 
  church_admin.church_nm_en AS '영문교회명', 
  Ifnull(
    branchlist.church_nm, churchlist.church_nm
  ) AS '지교회명', 
  Ifnull(
    branch_admin.church_nm_en, church_admin.church_nm_en
  ) AS '영문지교회명', 
	IFNULL(
		geoBranch.country_nm_ko, geo.country_nm_ko
	) AS '선교국가', 
  pstaff.lifeno AS '생명번호', 
  Concat(
    pstaff.name_ko, 
    Ifnull(
      Concat(
        '(', 
        LEFT(REPLACE(title.title,'봉사',''), 1), 
        ')'
      ), 
      ''
    )
  ) AS '한글이름(직분)', 
  pstaff.name_en AS '영문이름', 
  Ifnull(
    IF(
      position.position LIKE '%관리자%' 
      OR position.position LIKE '%당%' 
      OR position.position LIKE '%동%', 
      position.position, 
      theological.level
    ), 
    position2.position2
  ) AS '직책', 
  position2.position2 AS '직책2', 
  pstaff.birthday AS '생년월일', 
  pstaff.nationality AS '국적', 
  pstaff.home AS '고향', 
  spouse.home AS '사모고향', 
  IF(
    pstaff.appo_ovs IS NULL, appoint.start_dt, 
    pstaff.appo_ovs
  ) AS '(해외)최초발령일', 
--  belong.start_dt AS '현당회발령일', 
  IF(`position`.`position` IN ( '당회장', '당회장대리','동역' ),
     CASE 
		  WHEN `position`.`position` = '동역'
        THEN IF(`position`.`start_dt` >= `belong`.`start_dt`, `position`.`start_dt`, `belong`.`start_dt`)
        ELSE `belong`.`start_dt`
     END
     , NULL) AS '현당회발령일',
  spouse.lifeno AS '배우자생번', 
  Concat(
    spouse.name_ko, 
    Ifnull(
      Concat(
        '(', 
        LEFT(title_spouse.title, 1), 
        ')'
      ), 
      ''
    )
  ) AS '사모한글이름(직분)', 
  spouse.name_en AS '사모영문이름', 
  IF(
    spouse.lifeno IS NOT NULL, spouseposition.position_spouse, 
    NULL
  ) AS '사모직책', 
  spouse.birthday AS '배우자 생년월일', 
  IF(
    Ifnull(
      position.position, theological.level
    ) LIKE '당%' 
    OR Ifnull(
      position.position, theological.level
    ) LIKE '동%', 
    NULL, 
    theological.level
  ) AS '생도기수', 
  pstaff.ovs_dept AS '관리부서', 
  spouse.nationality AS '사모국적', 
  pstaff.name_ko AS '한글이름', 
  spouse.name_ko AS '사모한글이름', 
  pstaff.education AS '교육', 
  spouse.education AS '사모교육', 
  title.title AS '직분', 
  title_spouse.title AS '사모직분', 
  branchlist.church_sid AS '지교회코드' 
FROM 
  op_system.db_pastoralstaff pstaff 
  LEFT JOIN op_system.db_title title ON(
    pstaff.lifeno = title.lifeno 
    AND Curdate() BETWEEN title.start_dt 
    AND title.end_dt
  ) 
  LEFT JOIN op_system.db_position position ON(
    pstaff.lifeno = position.lifeno 
    AND Curdate() BETWEEN position.start_dt 
    AND position.end_dt
  ) 
  LEFT JOIN op_system.db_position2 position2 ON(
    pstaff.lifeno = position2.lifeno 
    AND Curdate() BETWEEN position2.start_dt 
    AND position2.end_dt
  ) 
  LEFT JOIN op_system.db_transfer belong ON(
    pstaff.lifeno = belong.lifeno 
    AND Curdate() BETWEEN belong.start_dt 
    AND belong.end_dt
  ) 
  LEFT JOIN op_system.db_churchlist churchlist ON(
    belong.church_sid = churchlist.church_sid
  ) 
  LEFT JOIN op_system.db_branchleader belongbranch ON(
    pstaff.lifeno = belongbranch.lifeno 
    AND Curdate() BETWEEN belongbranch.start_dt 
    AND belongbranch.end_dt
  ) 
  LEFT JOIN (
    SELECT 
      *
    FROM 
      op_system.db_churchlist
  ) branchlist ON(
    belongbranch.church_sid = branchlist.church_sid
  ) 
  LEFT JOIN op_system.db_theological theological ON(
    pstaff.lifeno = theological.lifeno 
    AND Curdate() BETWEEN theological.start_dt 
    AND theological.end_dt
  ) 
  LEFT JOIN op_system.db_pastoralwife spouse ON(
    pstaff.lifeno = spouse.lifeno_spouse
  ) 
  LEFT JOIN op_system.a_position_spouse spouseposition ON(
    IF(
      position.position LIKE '%관리자%' 
      OR position.position LIKE '%당%' 
      OR position.position LIKE '%동%', 
      position.position, 
      theological.level
    ) = spouseposition.position
  ) 
  LEFT JOIN (
    SELECT 
      op_system.db_position.lifeno AS lifeno, 
      Min(op_system.db_position.start_dt) AS start_dt, 
      op_system.db_position.position AS Position 
    FROM 
      op_system.db_position 
    WHERE 
      op_system.db_position.position = '당회장' 
      OR op_system.db_position.position = '당회장대리' 
      OR op_system.db_position.position = '동역' 
    GROUP BY 
      op_system.db_position.lifeno
  ) appoint ON(appoint.lifeno = pstaff.lifeno) 
  LEFT JOIN op_system.db_title title_spouse ON(
    title_spouse.lifeno = spouse.lifeno 
    AND Curdate() BETWEEN title_spouse.start_dt 
    AND title_spouse.end_dt
  ) 
  LEFT JOIN op_system.a_churchlist_admin church_admin ON(
    churchlist.church_sid = church_admin.church_sid
  ) 
  LEFT JOIN op_system.a_churchlist_admin branch_admin ON(
    branchlist.church_sid = branch_admin.church_sid
  ) 
  LEFT JOIN op_system.db_geodata geo ON geo.geo_cd = churchlist.geo_cd 
  LEFT JOIN op_system.db_geodata geoBranch ON geoBranch.geo_cd = branchlist.geo_cd 
GROUP BY 
  pstaff.lifeno 
ORDER BY 
  churchlist.church_nm, 
  IF(
    pstaff.appo_ovs IS NULL, appoint.start_dt, 
    pstaff.appo_ovs
  ) ;

-- 뷰 op_system.v0_theological_history 구조 내보내기
DROP VIEW IF EXISTS `v0_theological_history`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v0_theological_history`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v0_theological_history` AS SELECT 
	t.theological_cd
	,t.LifeNo
	,p.name_ko
	,t.`Level`
	,CASE 
		WHEN t.End_dt = '9999-12-31'
		THEN CONCAT(t.`Level`, ' 진행중')
		WHEN t.Resign_dt IS NOT NULL
		THEN CONCAT(t.`Level`, ' 성도복귀')
		ELSE CONCAT(t.`Level`, ' 수료')
	END AS 'CUR_STATUS'
	,t.Start_dt
	,t.End_dt
	,t.Resign_dt
	,t.church_sid
	,c.church_nm
FROM op_system.db_pastoralstaff p
LEFT JOIN op_system.db_theological t
	ON p.lifeno = t.LifeNo
LEFT JOIN op_system.db_churchlist c
	ON t.church_sid = c.church_sid
WHERE
	(t.LifeNo, t.End_dt) IN (SELECT t.LifeNo, MAX(t.End_dt) FROM op_system.db_theological t GROUP BY t.LifeNo) ;

-- 뷰 op_system.v_atten_detail_churchlist 구조 내보내기
DROP VIEW IF EXISTS `v_atten_detail_churchlist`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v_atten_detail_churchlist`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v_atten_detail_churchlist` AS SELECT
	basic.church_sid_custom
	,basic.start_dt
	,basic.end_dt
	,esta.church_sid
	,churchlist.church_nm
	,churchlist.church_gb
	,churchlist.ovs_dept
FROM 
	(
		SELECT
			esta.church_sid_custom
			,MIN(esta.start_dt) AS start_dt
			,MAX(esta.end_dt) AS end_dt
		FROM op_system.db_history_church_establish esta
		GROUP BY esta.church_sid_custom
	) AS basic
LEFT JOIN op_system.db_history_church_establish esta
	ON basic.church_sid_custom = esta.church_sid_custom
		AND basic.end_dt = esta.end_dt
LEFT JOIN op_system.db_churchlist churchlist
	ON esta.church_sid = churchlist.church_sid
WHERE esta.church_sid LIKE '%MC%'
ORDER BY churchlist.sort_order ;

-- 뷰 op_system.v_churchlist_final 구조 내보내기
DROP VIEW IF EXISTS `v_churchlist_final`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v_churchlist_final`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v_churchlist_final` AS SELECT 
  DISTINCT a.church_sid_custom AS '교회커스텀코드', 
  a.church_sid AS '교회코드', 
  b.church_nm AS '교회명(ko)', 
  d.church_nm_en AS '교회명(en)', 
  b.church_gb AS '교회구분', 
  b.main_church_cd AS '본교회코드', 
  e.church_nm AS '본교회명', 
  IF(
    b.church_gb = 'MC', d.manager_nm, 
    f.manager
  ) AS '관리자', 
  IF(
    b.church_gb = 'MC', d.title, 
    f.title
  ) AS '관리자직분', 
  IF(
    b.church_gb = 'MC', d.position, 
    f.position
  ) AS '관리자직책', 
  d.geo_cd AS 'GEO코드', 
  b.ovs_dept AS '관리부서', 
  d.latitude AS '위도', 
  d.longitude AS '경도', 
  b.suspend AS '논리삭제', 
  b.sort_order AS '정렬순서' 
FROM 
  db_history_church_establish a 
  LEFT JOIN db_churchlist_custom b ON(
    a.church_sid = b.church_sid
  ) 
  LEFT JOIN a_churchlist_admin d ON(
    a.church_sid = d.church_sid
  ) 
  LEFT JOIN db_churchlist_custom e ON(
    b.main_church_cd = e.church_sid
  ) 
  LEFT JOIN a_branch_admin f ON(
    a.church_sid = f.church_sid
  ) 
ORDER BY 
  b.sort_order ;

-- 뷰 op_system.v_churchlist_nomatch 구조 내보내기
DROP VIEW IF EXISTS `v_churchlist_nomatch`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v_churchlist_nomatch`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v_churchlist_nomatch` AS SELECT 
  a.church_sid AS church_sid, 
  a.church_nm AS church_nm, 
  a.church_gb AS church_gb, 
  c.church_nm AS main_church, 
  a.start_dt AS start_dt, 
  REPLACE(
    a.end_dt, '9999-12-31', '현재'
  ) AS end_dt, 
  a.ovs_dept AS ovs_dept 
FROM 
  db_churchlist a 
  LEFT JOIN db_history_church_establish b ON(
    a.church_sid = b.church_sid
  ) 
  LEFT JOIN db_churchlist c ON(
    a.main_church_cd = c.church_sid
  ) 
WHERE 
  b.church_esta_cd IS NULL ;

-- 뷰 op_system.v_familyinfo 구조 내보내기
DROP VIEW IF EXISTS `v_familyinfo`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v_familyinfo`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v_familyinfo` AS SELECT 
  family.family_id AS family_id, 
  family.family_cd AS family_cd, 
  IF(
    family.suspend = 0, 
    family.relations, 
    Concat(
      family.relations, '(별세)'
    )
  ) AS relations, 
  family.lifeno AS lifeno, 
  Ifnull(
    IF(
      family.lifeno = '', family.name_ko, 
      pstaff.name_ko
    ), 
    wife.name_ko
  ) AS name_ko, 
  Ifnull(
    IF(
      family.lifeno = '', family.name_en, 
      pstaff.name_en
    ), 
    wife.name_en
  ) AS name_en, 
  family.church_sid AS church_sid, 
  CASE WHEN Locate(
    ' ', 
    Ifnull(
      IF(
        family.lifeno = '', churchlist.church_nm, 
        churchlist_pstaff.church_nm
      ), 
      pinfo_spouse.교회명
    )
  ) > 0 THEN Reverse(
    LEFT(
      Reverse(
        Ifnull(
          IF(
            family.lifeno = '', churchlist.church_nm, 
            churchlist_pstaff.church_nm
          ), 
          pinfo_spouse.교회명
        )
      ), 
      Locate(
        ' ', 
        Reverse(
          Ifnull(
            IF(
              family.lifeno = '', churchlist.church_nm, 
              churchlist_pstaff.church_nm
            ), 
            pinfo_spouse.교회명
          )
        )
      ) -1
    )
  ) ELSE Ifnull(
    IF(
      family.lifeno = '', churchlist.church_nm, 
      churchlist_pstaff.church_nm
    ), 
    pinfo_spouse.교회명
  ) end AS church_nm, 
  IF(
    family.lifeno = '', 
    family.title, 
    IF(
      Substr(family.lifeno, 12, 1) = '1', 
      pinfo.직분, 
      pinfo_spouse.사모직분
    )
  ) AS title, 
  CASE WHEN family.lifeno = '' THEN family.position ELSE CASE WHEN Substr(family.lifeno, 12, 1) = '1' THEN Ifnull(
    pinfo.직책, pinfo.직책2
  ) ELSE Ifnull(
    pinfo_spouse.사모직책, position2_spouse.position2
  ) end end AS position, 
  REPLACE(Ifnull(
    IF(
      family.lifeno = '', family.birthday, 
      pstaff.birthday
    ), 
    wife.birthday
  ), '1900-01-01', '') AS birthday, 
  Ifnull(
    IF(
      family.lifeno = '', family.education, 
      pstaff.education
    ), 
    wife.education
  ) AS education, 
  IF(
    family.lifeno <> '', '본교성도', 
    family.religion
  ) AS religion, 
  family.recognition AS recognition, 
  family.memo AS memo, 
  family.suspend AS suspend, 
  Ifnull(
    IF(
      family.lifeno = '', churchlist.church_nm, 
      churchlist_pstaff.church_nm
    ), 
    pinfo_spouse.교회명
  ) AS churchFullName 
FROM 
  op_system.db_familyinfo family 
  LEFT JOIN op_system.db_pastoralstaff pstaff ON(
    family.lifeno = pstaff.lifeno
  ) 
  LEFT JOIN op_system.db_pastoralwife wife ON(
    family.lifeno = wife.lifeno
  ) 
  LEFT JOIN op_system.db_transfer trans ON(
    family.lifeno = trans.lifeno 
    AND Curdate() BETWEEN trans.start_dt 
    AND trans.end_dt
  ) 
  LEFT JOIN op_system.v0_pstaff_information_all pinfo ON(
    family.lifeno = pinfo.생명번호
  ) 
  LEFT JOIN op_system.v0_pstaff_information_all pinfo_spouse ON(
    family.lifeno = pinfo_spouse.배우자생번
  ) 
  LEFT JOIN op_system.db_position2 position2_spouse ON(
    family.lifeno = position2_spouse.lifeno 
    AND Curdate() BETWEEN position2_spouse.start_dt 
    AND position2_spouse.end_dt
  ) 
  LEFT JOIN op_system.db_theological theo ON(
    theo.lifeno <> 0 
    AND pstaff.lifeno <> 0 
    AND Curdate() BETWEEN theo.start_dt 
    AND theo.end_dt
  ) 
  LEFT JOIN op_system.db_churchlist churchlist ON(
    family.church_sid = churchlist.church_sid
  ) 
  LEFT JOIN op_system.db_churchlist churchlist_pstaff ON(
    trans.church_sid = churchlist_pstaff.church_sid
  ) ;

-- 뷰 op_system.v_history_church 구조 내보내기
DROP VIEW IF EXISTS `v_history_church`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v_history_church`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v_history_church` AS SELECT 
  DISTINCT temp.커스텀코드 AS '커스텀코드', 
  temp.교회코드 AS '교회코드', 
  temp.시작일 AS '날짜', 
  NULL AS '생명번호', 
  ifnull(
    concat(
      replace(
        replace(
          replace(
            temp.교회구분, 'MC', '정교회'
          ), 
          'PBC', 
          '예배소'
        ), 
        'BC', 
        '지교회'
      ), 
      ' 설립 / 관리교회명: ', 
      temp.관리교회명
    ), 
    '교회설립'
  ) AS '교회연혁' 
FROM 
  op_system.v0_history_church_temp temp 
WHERE 
  (
    temp.커스텀코드, temp.시작일
  ) IN (
    SELECT 
      a.커스텀코드, 
      min(a.시작일) 
    FROM 
      op_system.v0_history_church_temp a 
    GROUP BY 
      a.커스텀코드
  ) 
UNION 
SELECT 
  DISTINCT esta.church_sid_custom AS '커스텀코드', 
  temp.교회코드 AS '교회코드', 
  temp.시작일 AS '날짜', 
  NULL AS '생명번호', 
  concat(
    IF(
      Locate('_', temp.교회명) > 0, 
      LEFT(
        temp.교회명, 
        Locate('_', temp.교회명) -1
      ), 
      temp.교회명
    ), 
    ' ', 
    replace(
      replace(
        temp.교회구분, 'PBC', '예배소'
      ), 
      'BC', 
      '지교회'
    ), 
    ' 신규설립'
  ) AS '교회연혁' 
FROM 
    op_system.v0_history_church_temp temp 
    LEFT JOIN op_system.db_history_church_establish esta ON (
      temp.관리교회코드 = esta.church_sid
    )
WHERE 
  (
    temp.커스텀코드, temp.시작일
  ) IN (
    SELECT 
      a.커스텀코드, 
      min(a.시작일) 
    FROM 
      op_system.v0_history_church_temp a 
    GROUP BY 
      a.커스텀코드
  ) 
  AND temp.교회구분 NOT IN ('MC', 'HBC', 'PBC') 
UNION 
SELECT 
  temp2.커스텀코드 AS '커스텀코드', 
  temp2.교회코드 AS '교회코드', 
  temp2.날짜 AS '날짜', 
  temp2.생명번호 AS '생명번호', 
  temp2.교회연혁 AS '교회연혁' 
FROM 
  (
    SELECT 
      temp.커스텀코드 AS '커스텀코드', 
      temp.교회코드 AS '교회코드', 
      temp.선임일 AS '날짜', 
      temp.생명번호 AS '생명번호', 
      IF(
        temp.생명번호 = lag(temp.생명번호, 1) OVER (
          partition BY temp.커스텀코드 
          ORDER BY 
            temp.선임일
        ), 
        NULL, 
        concat(
          IF(
            temp.선임일 <> row1.최초선임일, 
            '관리자변경: ', '관리자선임: '
          ), 
          temp.한글이름, 
          '(', 
          LEFT(temp.직분, 1), 
          ')'
        )
      ) AS '교회연혁' 
    FROM 
      (
        op_system.v0_history_church_temp temp 
        LEFT JOIN (
          SELECT 
            a.커스텀코드 AS '커스텀코드', 
            min(a.선임일) AS '최초선임일' 
          FROM 
            op_system.v0_history_church_temp a 
          GROUP BY 
            a.커스텀코드
        ) row1 ON (
          row1.커스텀코드 = temp.커스텀코드
        )
      ) 
    WHERE 
      temp.선임일 IS NOT NULL
  ) temp2 
WHERE 
  temp2.교회연혁 IS NOT NULL 
UNION 
SELECT 
  DISTINCT temp.커스텀코드 AS '커스텀코드', 
  temp.교회코드 AS '교회코드', 
  temp.종료일 AS '날짜', 
  NULL AS '생명번호', 
  IF(
    temp.종료일 < curdate(), 
    concat(
      replace(
        replace(
          replace(
            temp.교회구분, 'MC', '정교회'
          ), 
          'PBC', 
          '예배소'
        ), 
        'BC', 
        '지교회'
      ), 
      ' 폐쇄'
    ), 
    ''
  ) AS '교회연혁' 
FROM 
  op_system.v0_history_church_temp temp 
WHERE 
  (
    temp.커스텀코드, temp.종료일
  ) IN (
    SELECT 
      a.커스텀코드, 
      max(a.종료일) 
    FROM 
      op_system.v0_history_church_temp a 
    GROUP BY 
      a.커스텀코드
  ) 
  AND IF(
    temp.종료일 < curdate(), 
    concat(
      replace(
        replace(
          replace(
            temp.교회구분, 'MC', '정교회'
          ), 
          'PBC', 
          '예배소'
        ), 
        'BC', 
        '지교회'
      ), 
      ' 폐쇄'
    ), 
    ''
  ) <> '' 
UNION 
SELECT 
  a.커스텀코드 AS '커스텀코드', 
  a.교회코드 AS '교회코드', 
  a.시작일 AS '날짜', 
  a.생명번호 AS '생명번호', 
  IF(
    a.교회연혁2 = '', 
    a.교회연혁1, 
    IF(
      a.교회연혁1 = '', 
      a.교회연혁2, 
      concat(
        a.교회연혁1, ' / ', a.교회연혁2
      )
    )
  ) AS 교회연혁 
FROM 
  (
    SELECT 
      temp.커스텀코드 AS '커스텀코드', 
      temp.교회코드 AS '교회코드', 
      temp.교회명 AS '교회명', 
      temp.교회구분 AS '교회구분', 
      lag(temp.교회구분, 1) OVER (
        partition BY temp.커스텀코드 
        ORDER BY 
          temp.시작일
      ) AS '직전교회구분', 
      temp.시작일 AS '시작일', 
      temp.생명번호 AS '생명번호', 
      IF(
        temp.교회구분 <> lag(temp.교회구분, 1) OVER (
          partition BY temp.커스텀코드 
          ORDER BY 
            temp.시작일
        ), 
        concat(
          replace(
            replace(
              replace(
                temp.교회구분, 'MC', '정교회'
              ), 
              'PBC', 
              '예배소'
            ), 
            'BC', 
            '지교회'
          ), 
          ' ', 
          IF(
            replace(
              replace(
                replace(temp.교회구분, 'MC', 1), 
                'PBC', 
                3
              ), 
              'BC', 
              2
            ) < replace(
              replace(
                replace(
                  lag(temp.교회구분, 1) OVER (
                    partition BY temp.커스텀코드 
                    ORDER BY 
                      temp.시작일
                  ), 
                  'MC', 
                  1
                ), 
                'PBC', 
                3
              ), 
              'BC', 
              2
            ), 
            '승격', 
            IF(
              replace(
                replace(
                  replace(
                    lag(temp.교회구분, 1) OVER (
                      partition BY temp.커스텀코드 
                      ORDER BY 
                        temp.시작일
                    ), 
                    'MC', 
                    1
                  ), 
                  'PBC', 
                  3
                ), 
                'BC', 
                2
              ) = 1, 
              concat(
                '하향조정 및 통합 / 관리교회: ', 
                temp.관리교회명
              ), 
              '하향조정'
            )
          )
        ), 
        ''
      ) AS '교회연혁1', 
      IF(
        IF(
          locate('_', temp.관리교회명) = 0, 
          temp.관리교회명, 
          LEFT(
            temp.관리교회명, 
            locate('_', temp.관리교회명) -1
          )
        ) <> IF(
          locate(
            '_', 
            lag(temp.관리교회명, 1) OVER (
              partition BY temp.커스텀코드 
              ORDER BY 
                temp.시작일
            )
          ) = 0, 
          lag(temp.관리교회명, 1) OVER (
            partition BY temp.커스텀코드 
            ORDER BY 
              temp.시작일
          ), 
          LEFT(
            lag(temp.관리교회명, 1) OVER (
              partition BY temp.커스텀코드 
              ORDER BY 
                temp.시작일
            ), 
            locate(
              '_', 
              lag(temp.관리교회명, 1) OVER (
                partition BY temp.커스텀코드 
                ORDER BY 
                  temp.시작일
              )
            ) -1
          )
        ), 
        concat(
          '관리교회변경: ', temp.관리교회명
        ), 
        ''
      ) AS '교회연혁2' 
    FROM 
      op_system.v0_history_church_temp temp
  ) a 
WHERE 
  IF(
    a.교회연혁2 = '', 
    a.교회연혁1, 
    IF(
      a.교회연혁1 = '', 
      a.교회연혁2, 
      concat(
        a.교회연혁1, ' / ', a.교회연혁2
      )
    )
  ) <> '' 
UNION 
SELECT 
  a.커스텀코드 AS '커스텀코드', 
  a.교회코드 AS '교회코드', 
  a.날짜 AS '날짜', 
  a.생명번호 AS '생명번호', 
  a.교회연혁 AS '교회연혁' 
FROM 
  (
    SELECT 
      lag(esta.church_sid_custom, 1) OVER (
        partition BY temp.커스텀코드 
        ORDER BY 
          temp.시작일
      ) AS '커스텀코드', 
      lag(temp.관리교회코드, 1) OVER (
        partition BY temp.커스텀코드 
        ORDER BY 
          temp.시작일
      ) AS '교회코드', 
      temp.시작일 AS 날짜, 
      NULL AS '생명번호', 
      IF(
        ifnull(
          lag(temp.관리교회명, 1) OVER (
            partition BY temp.커스텀코드 
            ORDER BY 
              temp.시작일
          ), 
          ''
        ) <> ifnull(temp.관리교회명, ''), 
        IF(
          temp.교회구분 = 'MC', 
          concat(
            IF(
              locate('_', temp.교회명) > 0, 
              reverse(
                LEFT(
                  reverse(
                    LEFT(
                      temp.교회명, 
                      locate('_', temp.교회명) -1
                    )
                  ), 
                  locate(
                    ' ', 
                    reverse(
                      LEFT(
                        temp.교회명, 
                        locate('_', temp.교회명) -1
                      )
                    )
                  ) -1
                )
              ), 
              reverse(
                LEFT(
                  reverse(temp.교회명), 
                  locate(
                    ' ', 
                    reverse(temp.교회명)
                  ) -1
                )
              )
            ), 
            ' 교회 분가'
          ), 
          concat(
            IF(
              locate('_', temp.교회명) > 0, 
              reverse(
                LEFT(
                  reverse(
                    LEFT(
                      temp.교회명, 
                      locate('_', temp.교회명) -1
                    )
                  ), 
                  locate(
                    ' ', 
                    reverse(
                      LEFT(
                        temp.교회명, 
                        locate('_', temp.교회명) -1
                      )
                    )
                  ) -1
                )
              ), 
              reverse(
                LEFT(
                  reverse(temp.교회명), 
                  locate(
                    ' ', 
                    reverse(temp.교회명)
                  ) -1
                )
              )
            ), 
            ' ', 
            replace(
              replace(
                temp.교회구분, 'PBC', '예배소'
              ), 
              'BC', 
              '지교회'
            ), 
            ' ', 
            IF(
              locate('_', temp.관리교회명) > 0, 
              IF(
                locate(' ', temp.관리교회명) = 0, 
                LEFT(
                  temp.관리교회명, 
                  locate('_', temp.관리교회명) -1
                ), 
                reverse(
                  LEFT(
                    reverse(
                      LEFT(
                        temp.관리교회명, 
                        locate('_', temp.관리교회명) -1
                      )
                    ), 
                    locate(
                      ' ', 
                      reverse(
                        LEFT(
                          temp.관리교회명, 
                          locate('_', temp.관리교회명) -1
                        )
                      )
                    ) -1
                  )
                )
              ), 
              IF(
                locate(' ', temp.관리교회명) = 0, 
                temp.관리교회명, 
                reverse(
                  LEFT(
                    reverse(temp.관리교회명), 
                    locate(
                      ' ', 
                      reverse(temp.관리교회명)
                    ) -1
                  )
                )
              )
            ), 
            ' 교회로 이관'
          )
        ), 
        ''
      ) AS '교회연혁' 
    FROM 
      (
        op_system.v0_history_church_temp temp 
        LEFT JOIN op_system.db_history_church_establish esta ON (
          esta.church_sid = temp.관리교회코드
        )
      ) 
    ORDER BY 
      temp.커스텀코드, 
      temp.시작일
  ) a 
WHERE 
  a.교회연혁 <> '' 
  AND a.커스텀코드 IS NOT NULL 
UNION 
SELECT 
  a.관리교회커스텀코드 AS '커스텀코드', 
  a.관리교회코드 AS '교회코드', 
  a.시작일 AS '날짜', 
  a.생명번호 AS '생명번호', 
  concat(
    IF(
      locate('_', a.교회명) > 0, 
      reverse(
        LEFT(
          reverse(
            LEFT(
              a.교회명, 
              locate('_', a.교회명) -1
            )
          ), 
          locate(
            ' ', 
            reverse(
              LEFT(
                a.교회명, 
                locate('_', a.교회명) -1
              )
            )
          ) -1
        )
      ), 
      reverse(
        LEFT(
          reverse(a.교회명), 
          locate(
            ' ', 
            reverse(a.교회명)
          ) -1
        )
      )
    ), 
    ' ', 
    replace(
      replace(
        replace(
          a.직전교회구분, 'PBC', 
          '예배소'
        ), 
        'BC', 
        '지교회'
      ), 
      'MC', 
      '교회'
    ), 
    ' ', 
    replace(
      replace(
        a.교회구분, 'PBC', '예배소'
      ), 
      'BC', 
      '지교회'
    ), 
    '로 ', 
    IF(
      a.교회연혁1 LIKE '%통합%', 
      '통합', '편입'
    )
  ) AS '교회연혁' 
FROM 
  (
    SELECT 
      temp.커스텀코드 AS '커스텀코드', 
      temp.교회코드 AS '교회코드', 
      temp.교회명 AS '교회명', 
      temp.교회구분 AS '교회구분', 
      lag(temp.교회구분, 1) OVER (
        partition BY temp.커스텀코드 
        ORDER BY 
          temp.시작일
      ) AS '직전교회구분', 
      temp.시작일 AS '시작일', 
      temp.생명번호 AS '생명번호', 
      esta.church_sid_custom AS '관리교회커스텀코드', 
      temp.관리교회코드 AS '관리교회코드', 
      temp.관리교회명 AS '관리교회명', 
      IF(
        temp.교회구분 <> lag(temp.교회구분, 1) OVER (
          partition BY temp.커스텀코드 
          ORDER BY 
            temp.시작일
        ), 
        concat(
          replace(
            replace(
              replace(
                temp.교회구분, 'MC', '정교회'
              ), 
              'PBC', 
              '예배소'
            ), 
            'BC', 
            '지교회'
          ), 
          ' ', 
          IF(
            replace(
              replace(
                replace(temp.교회구분, 'MC', 1), 
                'PBC', 
                3
              ), 
              'BC', 
              2
            ) < replace(
              replace(
                replace(
                  lag(temp.교회구분, 1) OVER (
                    partition BY temp.커스텀코드 
                    ORDER BY 
                      temp.시작일
                  ), 
                  'MC', 
                  1
                ), 
                'PBC', 
                3
              ), 
              'BC', 
              2
            ), 
            '승격', 
            IF(
              replace(
                replace(
                  replace(
                    lag(temp.교회구분, 1) OVER (
                      partition BY temp.커스텀코드 
                      ORDER BY 
                        temp.시작일
                    ), 
                    'MC', 
                    1
                  ), 
                  'PBC', 
                  3
                ), 
                'BC', 
                2
              ) = 1, 
              concat(
                '하향조정 및 통합 / 관리교회: ', 
                temp.관리교회명
              ), 
              '하향조정'
            )
          )
        ), 
        ''
      ) AS '교회연혁1', 
      IF(
        IF(
          locate('_', temp.관리교회명) = 0, 
          temp.관리교회명, 
          LEFT(
            temp.관리교회명, 
            locate('_', temp.관리교회명) -1
          )
        ) <> IF(
          locate(
            '_', 
            lag(temp.관리교회명, 1) OVER (
              partition BY temp.커스텀코드 
              ORDER BY 
                temp.시작일
            )
          ) = 0, 
          lag(temp.관리교회명, 1) OVER (
            partition BY temp.커스텀코드 
            ORDER BY 
              temp.시작일
          ), 
          LEFT(
            lag(temp.관리교회명, 1) OVER (
              partition BY temp.커스텀코드 
              ORDER BY 
                temp.시작일
            ), 
            locate(
              '_', 
              lag(temp.관리교회명, 1) OVER (
                partition BY temp.커스텀코드 
                ORDER BY 
                  temp.시작일
              )
            ) -1
          )
        ), 
        concat(
          '관리교회변경: ', temp.관리교회명
        ), 
        ''
      ) AS '교회연혁2' 
    FROM 
      (
        op_system.v0_history_church_temp temp 
        LEFT JOIN op_system.db_history_church_establish esta ON (
          esta.church_sid = temp.관리교회코드
        )
      )
  ) a 
WHERE 
  a.교회연혁1 LIKE '%통합%' 
  OR a.교회연혁2 LIKE '%관리교회변경%' 
UNION 
SELECT 
  b.church_sid_custom AS '커스텀코드', 
  a.church_sid AS '교회코드', 
  a.his_dt AS '날짜', 
  NULL AS '생명번호', 
  a.history AS '교회연혁' 
FROM 
  (
    op_system.db_history_church a 
    LEFT JOIN op_system.db_history_church_establish b ON (
      a.church_sid = b.church_sid
    )
  ) 
ORDER BY 
  커스텀코드, 
  날짜 ;

-- 뷰 op_system.v_phone 구조 내보내기
DROP VIEW IF EXISTS `v_phone`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v_phone`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v_phone` AS SELECT 
  IFNULL(geoBranch.country_nm_ko, geo.country_nm_ko) AS '선교국가', 
  tdiffer.time_different AS '시차', 
  esta.church_sid_custom AS '교회코드', 
  IF(
    branchlist.church_nm IS NULL, 
    churchlist.church_nm, branchlist.church_nm
  ) AS '교회명', 
  phone.phone AS '유선전화', 
  phone.wmcphone AS '인터넷전화', 
  pstaff.phone AS '선지자전화번호', 
  spouse.phone AS '배우자전화번호', 
  pstaff.lifeno AS '선지자생명번호', 
  Concat(
    pstaff.name_ko, 
    Ifnull(
      Concat(
        '(', 
        LEFT(title.title, 1), 
        ')'
      ), 
      ''
    )
  ) AS '한글이름(직분)', 
  Ifnull(
    IF(
      position.position LIKE '%관리자%' 
      OR position.position LIKE '%당%' 
      OR position.position LIKE '%동%', 
      position.position, 
      theological.level
    ), 
    position2.position2
  ) AS '직책', 
  spouse.lifeno AS '배우자생명번호', 
  Concat(
    spouse.name_ko, 
    Ifnull(
      Concat(
        '(', 
        LEFT(title_spouse.title, 1), 
        ')'
      ), 
      ''
    )
  ) AS '사모한글이름(직분)', 
  IF(
    spouse.lifeno IS NOT NULL, spouseposition.position_spouse, 
    NULL
  ) AS '사모직책', 
  churchlist.church_nm AS '관리교회명', 
  churchlist.ovs_dept AS '관리부서', 
  Concat(
    pstaff.name_en, 
    Ifnull(
      Concat(
        '(', 
        LEFT(title.title, 1), 
        ')'
      ), 
      ''
    )
  ) AS '영문이름', 
  Concat(
    spouse.name_en, 
    Ifnull(
      Concat(
        '(', 
        LEFT(title_spouse.title, 1), 
        ')'
      ), 
      ''
    )
  ) AS '사모영문이름', 
  phone.address AS '교회주소', 
  churchlist.church_sid AS '본교회코드', 
  Ifnull(
    branchlist.church_sid, churchlist.church_sid
  ) AS '지교회코드', 
  churchlist.sort_order AS '정렬순서', 
  pstaff.birthday AS '생년월일', 
  pinfo.`(해외)최초발령일` AS '최초발령일', 
  IF(
    branchlist_admin.church_nm_en IS NULL, 
    churchlist_admin.church_nm_en, 
    branchlist_admin.church_nm_en
  ) AS '영문교회명'
FROM 
  op_system.db_pastoralstaff pstaff 
  LEFT JOIN op_system.db_transfer belong ON(
    pstaff.lifeno = belong.lifeno 
    AND Curdate() BETWEEN belong.start_dt 
    AND belong.end_dt
  ) 
  LEFT JOIN op_system.db_churchlist churchlist ON(
    belong.church_sid = churchlist.church_sid
  ) 
  LEFT JOIN op_system.db_branchleader branch ON(
    pstaff.lifeno = branch.lifeno 
    AND Curdate() BETWEEN branch.start_dt 
    AND branch.end_dt
  ) 
  LEFT JOIN op_system.db_churchlist branchlist ON(
    branch.church_sid = branchlist.church_sid
  ) 
  LEFT JOIN op_system.a_churchlist_admin churchlist_admin ON churchlist.church_sid = churchlist_admin.church_sid 
  LEFT JOIN op_system.a_branch_admin branchlist_admin ON branchlist.church_sid = branchlist_admin.church_sid 
  LEFT JOIN op_system.db_pastoralwife spouse ON(
    pstaff.lifeno = spouse.lifeno_spouse
  ) 
  LEFT JOIN op_system.a_churchlist_admin church_admin ON(
    churchlist.church_sid = church_admin.church_sid
  ) 
  LEFT JOIN op_system.db_geodata geo 
  	ON churchlist.geo_cd = geo.geo_cd
  LEFT JOIN op_system.db_geodata geoBranch 
  	ON branchlist.geo_cd = geoBranch.geo_cd
  LEFT JOIN op_system.db_time_different tdiffer ON(
    tdiffer.country = IFNULL(geoBranch.country_nm_ko, geo.country_nm_ko)
  ) 
  LEFT JOIN op_system.db_history_church_establish esta ON(
    esta.church_sid = IF(
      branch.church_sid IS NULL, belong.church_sid, 
      branch.church_sid
    )
  ) 
  LEFT JOIN op_system.db_phone phone ON(
    esta.church_sid_custom = phone.church_sid
  ) 
  LEFT JOIN op_system.db_title title ON(
    pstaff.lifeno = title.lifeno 
    AND Curdate() BETWEEN title.start_dt 
    AND title.end_dt
  ) 
  LEFT JOIN op_system.db_position position ON(
    pstaff.lifeno = position.lifeno 
    AND Curdate() BETWEEN position.start_dt 
    AND position.end_dt
  ) 
  LEFT JOIN op_system.db_position2 position2 ON(
    pstaff.lifeno = position2.lifeno 
    AND Curdate() BETWEEN position2.start_dt 
    AND position2.end_dt
  ) 
  LEFT JOIN (
    SELECT 
      op_system.db_title.title_cd AS title_cd, 
      op_system.db_title.lifeno AS LifeNo, 
      op_system.db_title.start_dt AS Start_dt, 
      op_system.db_title.end_dt AS End_dt, 
      op_system.db_title.title AS Title 
    FROM 
      op_system.db_title
  ) title_spouse ON(
    title_spouse.lifeno = spouse.lifeno 
    AND Curdate() BETWEEN title_spouse.start_dt 
    AND title_spouse.end_dt
  ) 
  LEFT JOIN op_system.db_theological theological ON(
    pstaff.lifeno = theological.lifeno 
    AND Curdate() BETWEEN theological.start_dt 
    AND theological.end_dt
  ) 
  LEFT JOIN op_system.a_position_spouse spouseposition ON(
    Ifnull(
      position.position, theological.level
    ) = spouseposition.position
  ) 
  LEFT JOIN op_system.v0_pstaff_information pinfo ON(
    pstaff.lifeno = pinfo.생명번호
  ) 
ORDER BY 
  churchlist.sort_order IS NULL, 
  churchlist.sort_order, 
  pinfo.직책 IS NULL, 
  Field(
    pinfo.직책, '당회장', '당회장대리', 
    '당사모', '당대리사모', 
    '동역', '동사모', '지교회관리자', 
    '지관자사모', '예배소관리자', 
    '예관자사모', '예비생도1단계', 
    '예비생도2단계', '예비생도3단계', 
    '번역자', '행정직원', '자비량', 
    '건물관리'
  ), 
  Substr(
    pinfo.`한글이름(직분)`, 
    Locate(
      '(', pinfo.`한글이름(직분)`
    ), 
    1
  ) IS NULL, 
  Field(
    Substr(
      pinfo.`한글이름(직분)`, 
      Locate(
        '(', pinfo.`한글이름(직분)`
      ), 
      1
    ), 
    '목', 
    '장', 
    '전', 
    '집'
  ), 
  pinfo.`(해외)최초발령일`, 
  pstaff.birthday ;

-- 뷰 op_system.v_phone_export 구조 내보내기
DROP VIEW IF EXISTS `v_phone_export`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v_phone_export`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v_phone_export` AS SELECT 
  `a`.`선교국가` AS `선교국가`, 
  Date_format(`a`.`시차`, '%H:%i') AS `시차`, 
  `a`.`교회명` AS `교회명`, 
  `a`.`인터넷전화` AS `인터넷전화`, 
  `a`.`유선전화` AS `유선전화`, 
  `a`.`직책` AS `직책`, 
  `a`.`한글이름(직분)` AS `한글이름(직분)`, 
  `a`.`선지자전화번호` AS `선지자전화번호`, 
  `a`.`사모한글이름(직분)` AS `사모한글이름(직분)`, 
  `a`.`배우자전화번호` AS `배우자전화번호`, 
  `a`.`관리부서` AS `관리부서` 
FROM 
  `op_system`.`v_phone` `a` 
WHERE 
  `a`.`교회명` IS NOT NULL 
  AND `a`.`직책` IS NOT NULL 
  AND `a`.`정렬순서` >= (
    SELECT 
      `op_system`.`db_churchlist`.`sort_order` 
    FROM 
      `op_system`.`db_churchlist` 
    WHERE 
      `op_system`.`db_churchlist`.`church_nm` = '몽골 울란바토르'
  ) ;

-- 뷰 op_system.v_pstaff_detail 구조 내보내기
DROP VIEW IF EXISTS `v_pstaff_detail`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v_pstaff_detail`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v_pstaff_detail` AS SELECT 
  churchnm.church_nm AS '본교회명', 
  branchnm.church_nm AS '지교회명', 
  pstaff.name_ko AS '한글이름', 
  pstaff.name_en AS '영문이름', 
  title.title AS '직분', 
  IF(
    position1.position LIKE '%관리자%' 
    OR position1.position LIKE '%당%' 
    OR position1.position LIKE '%동%', 
    position1.position, 
    theological.level
  ) AS '직책', 
  position2.position2 AS '제2직책', 
  pstaff.birthday AS '생년월일', 
  pstaff.lifeno AS '생명번호', 
  pstaff.education AS '학력', 
  pstaff.baptism AS '침례권', 
  pstaff.ordination_prayer AS '안수일', 
  IF(
    position1.position IN (
      '당회장', '당회장대리', '동역'
    ), 
    '지원', 
    IF(
      pstaff.salary > 0, '지원', '미지원'
    )
  ) AS '유급', 
  pstaff.health AS '건강', 
  pstaff.appo_ovs AS '최초출국일', 
  pstaff.nationality AS '국적', 
  visa.visa AS '비자', 
  visa.end_dt AS '비자만료일', 
  pstaff.home AS '본가위치', 
  pstaff.family AS '가족사항', 
  wife.name_ko AS '사모한글이름', 
  wife.name_en AS '사모영문이름', 
  title_spouse.title AS '사모직분', 
  position_spouse.position_spouse AS '사모직책', 
  wife.birthday AS '사모생년월일', 
  wife.lifeno AS '사모생명번호', 
  wife.education AS '사모학력', 
  pstaff.lifeno_child1 AS '자녀1생명번호', 
  pstaff.birthday_child1 AS '자녀1생년월일', 
  pstaff.lifeno_child2 AS '자녀2생명번호', 
  pstaff.birthday_child2 AS '자녀2생년월일', 
  pstaff.lifeno_child3 AS '자녀3생명번호', 
  pstaff.birthday_child3 AS '자녀3생년월일', 
  wife.nationality AS '사모국적', 
  wife.health AS '사모건강', 
  visa_spouse.visa AS '사모비자', 
  visa_spouse.end_dt AS '사모비자만료일', 
  wife.home AS '친정위치', 
  wife.family AS '사모가족사항', 
  sermon.subject_count AS '발표주제', 
  sermon.score_avg AS '평균점수', 
  pstaff.theological_order AS '생도기수',
  theological.cur_status AS '예비생도단계',
  geo.country_nm_ko AS '선교국가',
  IF(`position1`.`position` IN ( '당회장', '당회장대리','동역' ),
     CASE 
		  WHEN `position1`.`position` = '동역'
        THEN IF(`position1`.`start_dt` >= `belong`.`start_dt`, `position1`.`start_dt`, `belong`.`start_dt`)
        ELSE `belong`.`start_dt`
     END
     , NULL) AS '현당회발령일'
FROM 
  db_pastoralstaff pstaff 
  LEFT JOIN db_transfer belong ON(
    pstaff.lifeno = belong.lifeno 
    AND CURDATE() BETWEEN belong.start_dt 
    AND belong.end_dt
  ) 
  LEFT JOIN db_churchlist churchnm ON(
    belong.church_sid = churchnm.church_sid
  ) 
  LEFT JOIN db_branchleader branchcd ON(
    pstaff.lifeno = branchcd.lifeno 
    AND CURDATE() BETWEEN branchcd.start_dt 
    AND branchcd.end_dt
  ) 
  LEFT JOIN db_churchlist branchnm ON(
    branchcd.church_sid = branchnm.church_sid
  ) 
  LEFT JOIN db_title title ON(
    pstaff.lifeno = title.lifeno 
    AND CURDATE() BETWEEN title.start_dt 
    AND title.end_dt
  ) 
  LEFT JOIN db_position position1 ON(
    pstaff.lifeno = position1.lifeno 
    AND CURDATE() BETWEEN position1.start_dt 
    AND position1.end_dt
  ) 
--  LEFT JOIN db_theological theological ON(
--    pstaff.lifeno = theological.lifeno 
--    AND CURDATE() BETWEEN theological.start_dt 
--    AND theological.end_dt
--  ) 
  LEFT JOIN v0_theological_history theological ON (
  	 pstaff.lifeno = theological.LifeNo
  )
  LEFT JOIN db_position2 position2 ON(
    pstaff.lifeno = position2.lifeno 
    AND CURDATE() BETWEEN position2.start_dt 
    AND position2.end_dt
  ) 
  LEFT JOIN db_visa visa ON(
    pstaff.lifeno = visa.lifeno 
    AND CURDATE() BETWEEN visa.start_dt 
    AND visa.end_dt
  ) 
  LEFT JOIN db_pastoralwife wife ON(
    pstaff.lifeno = wife.lifeno_spouse
  ) 
  LEFT JOIN db_title title_spouse ON(
    wife.lifeno = title_spouse.lifeno 
    AND CURDATE() BETWEEN title_spouse.start_dt 
    AND title_spouse.end_dt
  ) 
  LEFT JOIN a_position_spouse position_spouse ON(
    position1.position = position_spouse.position
  ) 
  LEFT JOIN db_visa visa_spouse ON(
    wife.lifeno = visa_spouse.lifeno 
    AND Curdate() BETWEEN visa_spouse.start_dt 
    AND visa_spouse.end_dt
  ) 
  LEFT JOIN db_sermon sermon ON(
    pstaff.lifeno = sermon.lifeno
  ) 
  LEFT JOIN op_system.db_geodata geo ON(
    IFNULL(branchnm.geo_cd, churchnm.geo_cd) = geo.geo_cd
  )
WHERE 
  (
    IF(
      position1.position LIKE '%관리자%' 
      OR position1.position LIKE '%당%' 
      OR position1.position LIKE '%동%', 
      position1.position, 
      theological.level
    ) NOT LIKE '%역장%' 
    AND IF(
      position1.position LIKE '%관리자%' 
      OR position1.position LIKE '%당%' 
      OR position1.position LIKE '%동%', 
      position1.position, 
      theological.level
    ) IS NOT NULL 
    OR position2.position2 IS NOT NULL
  ) 
  AND churchnm.church_nm IS NOT NULL 
  AND pstaff.suspend = 0 ;

-- 뷰 op_system.v_pstaff_detail_accomplishment 구조 내보내기
DROP VIEW IF EXISTS `v_pstaff_detail_accomplishment`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v_pstaff_detail_accomplishment`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v_pstaff_detail_accomplishment` AS SELECT 
  bcleader.church_sid AS '교회코드', 
  churchnm.church_nm AS '교회명', 
  bcleader.lifeno AS '생명번호', 
  pstaff.name_ko AS '한글이름', 
  atten.attendance_dt AS '날짜', 
  atten.once_all AS '전체1회', 
  atten.forth_all AS '전체4회', 
  atten.once_stu AS '학생1회', 
  atten.forth_stu AS '학생4회', 
  atten.tithe_stu AS '반차', 
  atten.baptism_all AS '침례', 
  atten.evangelist AS '전도인', 
  atten.ul AS '구역장', 
  atten.gl AS '지역장', 
  title.title AS '직분', 
  position1.position AS '직책', 
  bcleader.start_dt AS '관리시작일', 
  bcleader.end_dt AS '관리종료일', 
  churchnm.church_gb AS '교회구분' 
FROM 
  db_branchleader bcleader 
  LEFT JOIN db_churchlist churchnm ON(
    bcleader.church_sid = churchnm.church_sid
  ) 
  LEFT JOIN db_pastoralstaff pstaff ON(
    bcleader.lifeno = pstaff.lifeno
  ) 
  LEFT JOIN db_attendance atten ON(
    bcleader.church_sid = atten.church_sid 
    AND atten.attendance_dt BETWEEN Last_day(
      bcleader.start_dt + INTERVAL -1 month
    ) + INTERVAL 1 day 
    AND bcleader.end_dt + INTERVAL -1 month
  ) 
  LEFT JOIN db_title title ON(
    title.lifeno = bcleader.lifeno 
    AND atten.attendance_dt BETWEEN title.start_dt 
    AND title.end_dt
  ) 
  LEFT JOIN db_position position1 ON(
    position1.lifeno = bcleader.lifeno 
    AND atten.attendance_dt BETWEEN position1.start_dt 
    AND position1.end_dt
  ) 
WHERE 
  atten.attendance_dt IS NOT NULL 
UNION 
SELECT 
  transfer.church_sid AS '교회코드', 
  IF(
    churchnm.church_gb = 'MC', 
    Concat(
      churchnm.church_nm, ' 전체'
    ), 
    churchnm.church_nm
  ) AS '교회명', 
  pstaff.lifeno AS '생명번호', 
  pstaff.name_ko AS '한글이름', 
  atten.attendance_dt AS '날짜', 
  atten.once_all AS '전체1회', 
  atten.forth_all AS '전체4회', 
  atten.once_stu AS '학생1회', 
  atten.forth_stu AS '학생4회', 
  atten.tithe_stu AS '반차', 
  atten.baptism_all AS '침례', 
  atten.evangelist AS '전도인', 
  atten.ul AS '구역장', 
  atten.gl AS '지역장', 
  title.title AS '직분', 
  position1.position AS '직책', 
  transfer.start_dt AS '관리시작일', 
  transfer.end_dt AS '관리종료일', 
  churchnm.church_gb AS '교회구분' 
FROM 
  db_transfer transfer 
  LEFT JOIN db_pastoralstaff pstaff ON(
    pstaff.lifeno = transfer.lifeno
  ) 
  LEFT JOIN db_churchlist churchnm ON(
    transfer.church_sid = churchnm.church_sid
  ) 
  LEFT JOIN db_attendance atten ON(
    transfer.church_sid = atten.church_sid 
    AND (atten.attendance_dt BETWEEN LAST_DAY(
      transfer.start_dt + INTERVAL -1 MONTH
    ) + INTERVAL 1 DAY 
    AND transfer.end_dt + INTERVAL -1 MONTH)
  ) 
  LEFT JOIN db_title title ON(
    pstaff.lifeno = title.lifeno 
    AND LAST_DAY(atten.attendance_dt) BETWEEN title.start_dt 
    AND title.end_dt
  ) 
  LEFT JOIN db_position position1 ON(
    pstaff.lifeno = position1.lifeno 
    AND LAST_DAY(atten.attendance_dt) BETWEEN position1.start_dt 
    AND position1.end_dt
  ) 
WHERE 
  (
    (
      position1.position LIKE '%당%' 
      AND churchnm.church_gb IN ('MC', 'HBC')
    ) 
    OR (
      position1.position LIKE '%지교회관리자%' 
      -- 2023.09.25 유럽2과 오류접수: HBC가 MC로 승격한 경우 선지자상세정보 안나오는 문제 해결
	   -- 이 경우 지관지 직책을 유지하면서 본교회에 있는 분들 데이터가 잘못 나올 가능성이 있지만 그것은 엄밀히 말하면 직책 데이터가 잘못 입력되어 있는 경우이므로 아래와 같이 수정하는 것이 좋다고 판단됨
      AND churchnm.church_gb IN ('MC', 'HBC')
    )
  ) 
  AND atten.attendance_dt IS NOT NULL 
  AND churchnm.ovs_dept <> ""
ORDER BY 
  생명번호,
  날짜,
--  관리종료일 DESC, 
  교회코드 ;

-- 뷰 op_system.v_pstaff_detail_accomplishment_both 구조 내보내기
DROP VIEW IF EXISTS `v_pstaff_detail_accomplishment_both`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v_pstaff_detail_accomplishment_both`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v_pstaff_detail_accomplishment_both` AS SELECT 
  bcleader.church_sid AS '교회코드', 
  churchnm.church_nm AS '교회명', 
  bcleader.lifeno AS '생명번호', 
  pstaff.name_ko AS '한글이름', 
  atten.attendance_dt AS '날짜', 
  atten.once_all AS '전체1회', 
  atten.forth_all AS '전체4회', 
  atten.once_stu AS '학생1회', 
  atten.forth_stu AS '학생4회', 
  atten.tithe_stu AS '반차', 
  atten.baptism_all AS '침례', 
  atten.evangelist AS '전도인', 
  atten.ul AS '구역장', 
  atten.gl AS '지역장', 
  title.title AS '직분', 
  position1.position AS '직책', 
  bcleader.start_dt AS '관리시작일', 
  bcleader.end_dt AS '관리종료일', 
  churchnm.church_gb AS '교회구분' 
FROM 
  db_branchleader bcleader 
  LEFT JOIN db_churchlist churchnm ON(
    bcleader.church_sid = churchnm.church_sid
  ) 
  LEFT JOIN db_pastoralstaff pstaff ON(
    bcleader.lifeno = pstaff.lifeno
  ) 
  LEFT JOIN db_attendance atten ON(
    bcleader.church_sid = atten.church_sid 
    AND atten.attendance_dt BETWEEN Last_day(
      bcleader.start_dt + INTERVAL -1 month
    ) + INTERVAL 1 day 
    AND bcleader.end_dt + INTERVAL -1 month
  ) 
  LEFT JOIN db_title title ON(
    title.lifeno = bcleader.lifeno 
    AND atten.attendance_dt BETWEEN title.start_dt 
    AND title.end_dt
  ) 
  LEFT JOIN db_position position1 ON(
    position1.lifeno = bcleader.lifeno 
    AND atten.attendance_dt BETWEEN position1.start_dt 
    AND position1.end_dt
  ) 
WHERE 
  atten.attendance_dt IS NOT NULL 
UNION 
SELECT 
  transfer.church_sid AS '교회코드', 
  IF(
    churchnm.church_gb = 'MC', 
    Concat(
      churchnm.church_nm, ' 전체'
    ), 
    churchnm.church_nm
  ) AS '교회명', 
  pstaff.lifeno AS '생명번호', 
  pstaff.name_ko AS '한글이름', 
  atten.attendance_dt AS '날짜', 
  atten.once_all AS '전체1회', 
  atten.forth_all AS '전체4회', 
  atten.once_stu AS '학생1회', 
  atten.forth_stu AS '학생4회', 
  atten.tithe_stu AS '반차', 
  atten.baptism_all AS '침례', 
  atten.evangelist AS '전도인', 
  atten.ul AS '구역장', 
  atten.gl AS '지역장', 
  title.title AS '직분', 
  position1.position AS '직책', 
  transfer.start_dt AS '관리시작일', 
  transfer.end_dt AS '관리종료일', 
  churchnm.church_gb AS '교회구분' 
FROM 
  db_transfer transfer 
  LEFT JOIN db_pastoralstaff pstaff ON(
    pstaff.lifeno = transfer.lifeno
  ) 
  LEFT JOIN db_churchlist churchnm ON(
    transfer.church_sid = churchnm.church_sid
  ) 
  LEFT JOIN db_attendance atten ON(
    transfer.church_sid = atten.church_sid 
    AND atten.attendance_dt BETWEEN Last_day(
      transfer.start_dt + INTERVAL -1 month
    ) + INTERVAL 1 day 
    AND transfer.end_dt + INTERVAL -1 month
  ) 
  LEFT JOIN db_title title ON(
    pstaff.lifeno = title.lifeno 
    AND Last_day(atten.attendance_dt) BETWEEN title.start_dt 
    AND title.end_dt
  ) 
  LEFT JOIN db_position position1 ON(
    pstaff.lifeno = position1.lifeno 
    AND Last_day(atten.attendance_dt) BETWEEN position1.start_dt 
    AND position1.end_dt
  ) 
WHERE 
  position1.position LIKE '%당%' 
  AND atten.attendance_dt IS NOT NULL 
UNION 
SELECT 
  transfer.church_sid AS '교회코드', 
  churchnm.church_nm AS '교회명', 
  pstaff.lifeno AS '생명번호', 
  pstaff.name_ko AS '한글이름', 
  atten.attendance_dt AS '날짜', 
  atten.once_all AS '전체1회', 
  atten.forth_all AS '전체4회', 
  atten.once_stu AS '학생1회', 
  atten.forth_stu AS '학생4회', 
  atten.tithe_stu AS '반차', 
  atten.baptism_all AS '침례', 
  atten.evangelist AS '전도인', 
  atten.ul AS '구역장', 
  atten.gl AS '지역장', 
  title.title AS '직분', 
  position1.position AS '직책', 
  transfer.start_dt AS '관리시작일', 
  transfer.end_dt AS '관리종료일', 
  churchnm.church_gb AS '교회구분' 
FROM 
  db_transfer transfer 
  LEFT JOIN db_pastoralstaff pstaff ON(
    pstaff.lifeno = transfer.lifeno
  ) 
  LEFT JOIN db_churchlist_custom churchnm ON(
    REPLACE(
      transfer.church_sid, 'MC', 'MM'
    ) = churchnm.church_sid
  ) 
  LEFT JOIN db_attendance atten ON(
    REPLACE(
      transfer.church_sid, 'MC', 'MM'
    ) = atten.church_sid 
    AND atten.attendance_dt BETWEEN Last_day(
      transfer.start_dt + INTERVAL -1 month
    ) + INTERVAL 1 day 
    AND transfer.end_dt + INTERVAL -1 month
  ) 
  LEFT JOIN db_title title ON(
    pstaff.lifeno = title.lifeno 
    AND Last_day(atten.attendance_dt) BETWEEN title.start_dt 
    AND title.end_dt
  ) 
  LEFT JOIN db_position position1 ON(
    pstaff.lifeno = position1.lifeno 
    AND Last_day(atten.attendance_dt) BETWEEN position1.start_dt 
    AND position1.end_dt
  ) 
WHERE 
  position1.position LIKE '%당%' 
  AND atten.attendance_dt IS NOT NULL 
  AND churchnm.ovs_dept <> "" 
ORDER BY 
  관리종료일, 
  교회코드 DESC, 
  생명번호, 
  날짜 ;

-- 뷰 op_system.v_pstaff_detail_accomplishment_main 구조 내보내기
DROP VIEW IF EXISTS `v_pstaff_detail_accomplishment_main`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v_pstaff_detail_accomplishment_main`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v_pstaff_detail_accomplishment_main` AS SELECT 
  bcleader.church_sid AS '교회코드', 
  churchnm.church_nm AS '교회명', 
  bcleader.lifeno AS '생명번호', 
  pstaff.name_ko AS '한글이름', 
  atten.attendance_dt AS '날짜', 
  atten.once_all AS '전체1회', 
  atten.forth_all AS '전체4회', 
  atten.once_stu AS '학생1회', 
  atten.forth_stu AS '학생4회', 
  atten.tithe_stu AS '반차', 
  atten.baptism_all AS '침례', 
  atten.evangelist AS '전도인', 
  atten.ul AS '구역장', 
  atten.gl AS '지역장', 
  title.title AS '직분', 
  position1.position AS '직책', 
  bcleader.start_dt AS '관리시작일', 
  bcleader.end_dt AS '관리종료일', 
  churchnm.church_gb AS '교회구분' 
FROM 
  db_branchleader bcleader 
  LEFT JOIN db_churchlist churchnm ON(
    bcleader.church_sid = churchnm.church_sid
  ) 
  LEFT JOIN db_pastoralstaff pstaff ON(
    bcleader.lifeno = pstaff.lifeno
  ) 
  LEFT JOIN db_attendance atten ON(
    bcleader.church_sid = atten.church_sid 
    AND atten.attendance_dt BETWEEN Last_day(
      bcleader.start_dt + INTERVAL -1 month
    ) + INTERVAL 1 day 
    AND bcleader.end_dt + INTERVAL -1 month
  ) 
  LEFT JOIN db_title title ON(
    title.lifeno = bcleader.lifeno 
    AND atten.attendance_dt BETWEEN title.start_dt 
    AND title.end_dt
  ) 
  LEFT JOIN db_position position1 ON(
    position1.lifeno = bcleader.lifeno 
    AND atten.attendance_dt BETWEEN position1.start_dt 
    AND position1.end_dt
  ) 
WHERE 
  atten.attendance_dt IS NOT NULL 
UNION 
SELECT 
  transfer.church_sid AS '교회코드', 
  churchnm.church_nm AS '교회명', 
  pstaff.lifeno AS '생명번호', 
  pstaff.name_ko AS '한글이름', 
  atten.attendance_dt AS '날짜', 
  atten.once_all AS '전체1회', 
  atten.forth_all AS '전체4회', 
  atten.once_stu AS '학생1회', 
  atten.forth_stu AS '학생4회', 
  atten.tithe_stu AS '반차', 
  atten.baptism_all AS '침례', 
  atten.evangelist AS '전도인', 
  atten.ul AS '구역장', 
  atten.gl AS '지역장', 
  title.title AS '직분', 
  position1.position AS '직책', 
  transfer.start_dt AS '관리시작일', 
  transfer.end_dt AS '관리종료일', 
  churchnm.church_gb AS '교회구분' 
FROM 
  db_transfer transfer 
  LEFT JOIN db_pastoralstaff pstaff ON(
    pstaff.lifeno = transfer.lifeno
  ) 
  LEFT JOIN db_churchlist_custom churchnm ON(
    REPLACE(
      transfer.church_sid, 'MC', 'MM'
    ) = churchnm.church_sid
  ) 
  LEFT JOIN db_attendance atten ON(
    REPLACE(
      transfer.church_sid, 'MC', 'MM'
    ) = atten.church_sid 
    AND atten.attendance_dt BETWEEN Last_day(
      transfer.start_dt + INTERVAL -1 month
    ) + INTERVAL 1 day 
    AND transfer.end_dt + INTERVAL -1 month
  ) 
  LEFT JOIN db_title title ON(
    pstaff.lifeno = title.lifeno 
    AND Last_day(atten.attendance_dt) BETWEEN title.start_dt 
    AND title.end_dt
  ) 
  LEFT JOIN db_position position1 ON(
    pstaff.lifeno = position1.lifeno 
    AND Last_day(atten.attendance_dt) BETWEEN position1.start_dt 
    AND position1.end_dt
  ) 
WHERE 
  position1.position LIKE '%당%' 
  AND atten.attendance_dt IS NOT NULL 
  AND churchnm.ovs_dept <> "" 
ORDER BY 
  관리종료일, 
  교회코드 DESC, 
  생명번호, 
  날짜 ;

-- 뷰 op_system.v_pstaff_detail_concise_transfer_history 구조 내보내기
DROP VIEW IF EXISTS `v_pstaff_detail_concise_transfer_history`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v_pstaff_detail_concise_transfer_history`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v_pstaff_detail_concise_transfer_history` AS SELECT 
  churchnm.church_nm AS '교회명', 
  ADDDATE(
    LAST_DAY(
      bcleader.start_dt
    ), 
    INTERVAL 1 DAY
  ) AS '시작일', 
  ADDDATE(
    LAST_DAY(
      ADDDATE(
        IF(
          bcleader.End_dt = '9999-12-31', 
          CURDATE(), 
          bcleader.End_dt
        ), 
        INTERVAL -2 MONTH
      )
    ), 
    INTERVAL 1 DAY
  ) AS '종료일', 
  bcleader.start_dt AS '관리시작일', 
  bcleader.end_dt AS '관리종료일', 
  PERIOD_DIFF(
    DATE_FORMAT(
      IF(
        bcleader.End_dt = '9999-12-31', 
        CURDATE(), 
        bcleader.End_dt
      ), 
      '%Y%m'
    ), 
    DATE_FORMAT(bcleader.start_dt, '%Y%m')
  ) AS '기간', 
  title.title AS '직분', 
  position1.position AS '직책', 
  churchnm.church_gb AS '교회구분', 
  bcleader.lifeno AS '생명번호',
  bcleader.church_sid AS '교회코드'
FROM 
  db_branchleader bcleader 
  LEFT JOIN db_churchlist churchnm ON(
    bcleader.church_sid = churchnm.church_sid
  ) 
  LEFT JOIN db_pastoralstaff pstaff ON(bcleader.lifeno = pstaff.lifeno) 
  LEFT JOIN db_title title ON(
    title.lifeno = bcleader.lifeno 
    AND LAST_DAY(bcleader.Start_dt) BETWEEN title.start_dt 
    AND title.end_dt
  ) 
  LEFT JOIN db_position position1 ON(
    position1.lifeno = bcleader.lifeno 
    AND LAST_DAY(bcleader.Start_dt) BETWEEN position1.start_dt 
    AND position1.end_dt
  ) 
UNION 
SELECT 
  IF(
    churchnm.church_gb = 'MC', 
    Concat(churchnm.church_nm, ' 전체'), 
    churchnm.church_nm
  ) AS '교회명', 
  ADDDATE(
    LAST_DAY(
      IF(POSITION1.start_dt >= transfer.start_dt,
        POSITION1.start_dt,
		  transfer.start_dt)
    ), 
    INTERVAL 1 DAY
  ) AS '시작일', 
  ADDDATE(
    LAST_DAY(
      ADDDATE(
        IF(
          transfer.End_dt = '9999-12-31', 
          CURDATE(), 
          IF(POSITION1.end_dt <= transfer.end_dt,
          POSITION1.end_dt,
			 transfer.End_dt)
        ), 
        INTERVAL -2 MONTH
      )
    ), 
    INTERVAL 1 DAY
  ) AS '종료일', 
  transfer.start_dt AS '관리시작일', 
  transfer.end_dt AS '관리종료일', 
  PERIOD_DIFF(
    DATE_FORMAT(
      IF(
        transfer.End_dt = '9999-12-31', 
        CURDATE(), 
        IF(POSITION1.end_dt <= transfer.end_dt,
        POSITION1.end_dt,
		  transfer.End_dt)
      ), 
      '%Y%m'
    ), 
    DATE_FORMAT(transfer.start_dt, '%Y%m')
  ) AS '기간', 
  title.title AS '직분', 
  position1.position AS '직책', 
  churchnm.church_gb AS '교회구분', 
  pstaff.lifeno AS '생명번호',
  transfer.church_sid AS '교회코드'
FROM 
  db_transfer transfer 
  LEFT JOIN db_pastoralstaff pstaff ON(pstaff.lifeno = transfer.lifeno) 
  LEFT JOIN db_churchlist churchnm ON(
    transfer.church_sid = churchnm.church_sid
  ) 
  LEFT JOIN db_title title ON(
    pstaff.lifeno = title.lifeno 
    AND LAST_DAY(transfer.start_dt) BETWEEN title.start_dt 
    AND title.end_dt
  ) 
  LEFT JOIN db_position position1 ON(
    pstaff.lifeno = position1.lifeno 
    AND (
		 (position1.start_dt BETWEEN transfer.start_dt
		    AND transfer.end_dt)
	    OR (LAST_DAY(transfer.start_dt) BETWEEN position1.start_dt 
		    AND position1.end_dt
		 )
	 )
    
  )
WHERE 
  ((
    position1.position LIKE '%당%' 
    AND churchnm.church_gb IN ('MC', 'HBC')
  )
  OR (
    position1.position LIKE '%지교회관리자%' 
    -- 2023.09.25 유럽2과 오류접수: HBC가 MC로 승격한 경우 선지자상세정보 안나오는 문제 해결
	 -- 이 경우 지관지 직책을 유지하면서 본교회에 있는 분들 데이터가 잘못 나올 가능성이 있지만 그것은 엄밀히 말하면 직책 데이터가 잘못 입력되어 있는 경우이므로 아래와 같이 수정하는 것이 좋다고 판단됨
	 AND churchnm.church_gb IN ('MC', 'HBC') 
  ))
  AND churchnm.ovs_dept <> ""
ORDER BY 
  생명번호, 
  관리종료일 DESC --  교회코드 DESC, ;

-- 뷰 op_system.v_pstaff_detail_concise_transfer_history_both 구조 내보내기
DROP VIEW IF EXISTS `v_pstaff_detail_concise_transfer_history_both`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v_pstaff_detail_concise_transfer_history_both`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v_pstaff_detail_concise_transfer_history_both` AS SELECT 
  --  bcleader.church_sid AS '교회코드', 
  churchnm.church_nm AS '교회명', 
  --  pstaff.name_ko AS '한글이름', 
  ADDDATE(
    LAST_DAY(
      bcleader.start_dt
--		ADDDATE(
--        bcleader.start_dt, INTERVAL -1 MONTH
--      )
    ), 
    INTERVAL 1 DAY
  ) AS '시작일', 
  ADDDATE(
    LAST_DAY(
      ADDDATE(
        IF(
          bcleader.End_dt = '9999-12-31', 
          CURDATE(), 
          bcleader.End_dt
        ), 
        INTERVAL -2 MONTH
      )
    ), 
    INTERVAL 1 DAY
  ) AS '종료일', 
  bcleader.start_dt AS '관리시작일', 
  bcleader.end_dt AS '관리종료일', 
  PERIOD_DIFF(
    DATE_FORMAT(
      IF(
        bcleader.End_dt = '9999-12-31', 
        CURDATE(), 
        bcleader.End_dt
      ), 
      '%Y%m'
    ), 
    DATE_FORMAT(bcleader.start_dt, '%Y%m')
  ) AS '기간', 
  title.title AS '직분', 
  position1.position AS '직책', 
  churchnm.church_gb AS '교회구분', 
  bcleader.lifeno AS '생명번호',
  bcleader.church_sid AS '교회코드'
FROM 
  db_branchleader bcleader 
  LEFT JOIN db_churchlist churchnm ON(
    bcleader.church_sid = churchnm.church_sid
  ) 
  LEFT JOIN db_pastoralstaff pstaff ON(bcleader.lifeno = pstaff.lifeno) 
  LEFT JOIN db_title title ON(
    title.lifeno = bcleader.lifeno 
    AND LAST_DAY(bcleader.Start_dt) BETWEEN title.start_dt 
    AND title.end_dt
  ) 
  LEFT JOIN db_position position1 ON(
    position1.lifeno = bcleader.lifeno 
    AND LAST_DAY(bcleader.Start_dt) BETWEEN position1.start_dt 
    AND position1.end_dt
  ) 
UNION 
SELECT 
  --  transfer.church_sid AS '교회코드', 
  IF(
    churchnm.church_gb = 'MC', 
    Concat(churchnm.church_nm, ' 전체'), 
    churchnm.church_nm
  ) AS '교회명', 
  --  pstaff.name_ko AS '한글이름', 
  ADDDATE(
    LAST_DAY(
      IF(POSITION1.start_dt >= transfer.start_dt,
        POSITION1.start_dt,
		  transfer.start_dt)
--		ADDDATE(
--        IF(POSITION1.start_dt >= transfer.start_dt,
--        POSITION1.start_dt,
--		  transfer.start_dt), 
--		  INTERVAL -1 MONTH
--      )
    ), 
    INTERVAL 1 DAY
  ) AS '시작일', 
  ADDDATE(
    LAST_DAY(
      ADDDATE(
        IF(
          transfer.End_dt = '9999-12-31', 
          CURDATE(), 
          IF(POSITION1.end_dt <= transfer.end_dt,
          POSITION1.end_dt,
			 transfer.End_dt)
        ), 
        INTERVAL -2 MONTH
      )
    ), 
    INTERVAL 1 DAY
  ) AS '종료일', 
  transfer.start_dt AS '관리시작일', 
  transfer.end_dt AS '관리종료일', 
  PERIOD_DIFF(
    DATE_FORMAT(
      IF(
        transfer.End_dt = '9999-12-31', 
        CURDATE(), 
        IF(POSITION1.end_dt <= transfer.end_dt,
        POSITION1.end_dt,
		  transfer.End_dt)
      ), 
      '%Y%m'
    ), 
    DATE_FORMAT(transfer.start_dt, '%Y%m')
  ) AS '기간', 
  title.title AS '직분', 
  position1.position AS '직책', 
  churchnm.church_gb AS '교회구분', 
  pstaff.lifeno AS '생명번호',
  transfer.church_sid AS '교회코드'
FROM 
  db_transfer transfer 
  LEFT JOIN db_pastoralstaff pstaff ON(pstaff.lifeno = transfer.lifeno) 
  LEFT JOIN db_churchlist churchnm ON(
    transfer.church_sid = churchnm.church_sid
  ) 
  LEFT JOIN db_title title ON(
    pstaff.lifeno = title.lifeno 
    AND LAST_DAY(transfer.start_dt) BETWEEN title.start_dt 
    AND title.end_dt
  ) 
  LEFT JOIN db_position position1 ON(
    pstaff.lifeno = position1.lifeno 
    AND LAST_DAY(transfer.start_dt) BETWEEN position1.start_dt 
    AND position1.end_dt
  ) 
  
UNION 
SELECT 
  --  transfer.church_sid AS '교회코드', 
  IF(
    churchnm.church_gb = 'MC', 
    Concat(churchnm.church_nm, ' 전체'), 
    churchnm.church_nm
  ) AS '교회명', 
  --  pstaff.name_ko AS '한글이름', 
  ADDDATE(
    LAST_DAY(
      ADDDATE(
        IF(POSITION1.start_dt >= transfer.start_dt,
        POSITION1.start_dt,
		  transfer.start_dt), 
		  INTERVAL -1 MONTH
      )
    ), 
    INTERVAL 1 DAY
  ) AS '시작일', 
  ADDDATE(
    LAST_DAY(
      ADDDATE(
        IF(
          transfer.End_dt = '9999-12-31', 
          CURDATE(), 
          IF(POSITION1.end_dt <= transfer.end_dt,
          POSITION1.end_dt,
			 transfer.End_dt)
        ), 
        INTERVAL -2 MONTH
      )
    ), 
    INTERVAL 1 DAY
  ) AS '종료일', 
  transfer.start_dt AS '관리시작일', 
  transfer.end_dt AS '관리종료일', 
  PERIOD_DIFF(
    DATE_FORMAT(
      IF(
        transfer.End_dt = '9999-12-31', 
        CURDATE(), 
        IF(POSITION1.end_dt <= transfer.end_dt,
        POSITION1.end_dt,
		  transfer.End_dt)
      ), 
      '%Y%m'
    ), 
    DATE_FORMAT(transfer.start_dt, '%Y%m')
  ) AS '기간', 
  title.title AS '직분', 
  position1.position AS '직책', 
  churchnm.church_gb AS '교회구분', 
  pstaff.lifeno AS '생명번호',
  transfer.church_sid AS '교회코드'
FROM 
  db_transfer transfer 
  LEFT JOIN db_pastoralstaff pstaff ON(pstaff.lifeno = transfer.lifeno) 
  LEFT JOIN db_churchlist_custom churchnm ON(
    REPLACE(transfer.church_sid, 'MC', 'MM') = churchnm.church_sid
  ) 
  LEFT JOIN db_title title ON(
    pstaff.lifeno = title.lifeno 
    AND LAST_DAY(transfer.start_dt) BETWEEN title.start_dt 
    AND title.end_dt
  ) 
  LEFT JOIN db_position position1 ON(
    pstaff.lifeno = position1.lifeno 
    AND LAST_DAY(transfer.start_dt) BETWEEN position1.start_dt 
    AND position1.end_dt
  ) 
WHERE 
  ((
    position1.position LIKE '%당%' 
    AND churchnm.church_gb IN ('MC', 'HBC', 'MM')
  )
  OR (
    position1.position LIKE '%지교회관리자%' 
    -- 2023.09.25 유럽2과 오류접수: HBC가 MC로 승격한 경우 선지자상세정보 안나오는 문제 해결
	 -- 이 경우 지관지 직책을 유지하면서 본교회에 있는 분들 데이터가 잘못 나올 가능성이 있지만 그것은 엄밀히 말하면 직책 데이터가 잘못 입력되어 있는 경우이므로 아래와 같이 수정하는 것이 좋다고 판단됨
    AND churchnm.church_gb IN ('MC', 'HBC')
  ))
  AND churchnm.ovs_dept <> ""
ORDER BY 
  생명번호, 
  관리종료일 DESC --  교회코드 DESC, ;

-- 뷰 op_system.v_pstaff_detail_concise_transfer_history_main 구조 내보내기
DROP VIEW IF EXISTS `v_pstaff_detail_concise_transfer_history_main`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v_pstaff_detail_concise_transfer_history_main`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v_pstaff_detail_concise_transfer_history_main` AS SELECT 
  --  bcleader.church_sid AS '교회코드', 
  churchnm.church_nm AS '교회명', 
  --  pstaff.name_ko AS '한글이름', 
  ADDDATE(
    LAST_DAY(
      bcleader.start_dt
--		ADDDATE(
--        bcleader.start_dt, INTERVAL -1 MONTH
--      )
    ), 
    INTERVAL 1 DAY
  ) AS '시작일', 
  ADDDATE(
    LAST_DAY(
      ADDDATE(
        IF(
          bcleader.End_dt = '9999-12-31', 
          CURDATE(), 
          bcleader.End_dt
        ), 
        INTERVAL -2 MONTH
      )
    ), 
    INTERVAL 1 DAY
  ) AS '종료일', 
  bcleader.start_dt AS '관리시작일', 
  bcleader.end_dt AS '관리종료일', 
  PERIOD_DIFF(
    DATE_FORMAT(
      IF(
        bcleader.End_dt = '9999-12-31', 
        CURDATE(), 
        bcleader.End_dt
      ), 
      '%Y%m'
    ), 
    DATE_FORMAT(bcleader.start_dt, '%Y%m')
  ) AS '기간', 
  title.title AS '직분', 
  position1.position AS '직책', 
  churchnm.church_gb AS '교회구분', 
  bcleader.lifeno AS '생명번호',
  bcleader.church_sid AS '교회코드'
FROM 
  db_branchleader bcleader 
  LEFT JOIN db_churchlist_custom churchnm ON(
    bcleader.church_sid = churchnm.church_sid
  ) 
  LEFT JOIN db_pastoralstaff pstaff ON(bcleader.lifeno = pstaff.lifeno) 
  LEFT JOIN db_title title ON(
    title.lifeno = bcleader.lifeno 
    AND LAST_DAY(bcleader.Start_dt) BETWEEN title.start_dt 
    AND title.end_dt
  ) 
  LEFT JOIN db_position position1 ON(
    position1.lifeno = bcleader.lifeno 
    AND LAST_DAY(bcleader.Start_dt) BETWEEN position1.start_dt 
    AND position1.end_dt
  ) 
UNION 
SELECT 
  --  transfer.church_sid AS '교회코드', 
  IF(
    churchnm.church_gb = 'MC', 
    Concat(churchnm.church_nm, ' 전체'), 
    churchnm.church_nm
  ) AS '교회명', 
  --  pstaff.name_ko AS '한글이름', 
  ADDDATE(
    LAST_DAY(
      IF(POSITION1.start_dt >= transfer.start_dt,
        POSITION1.start_dt,
		  transfer.start_dt)
--		ADDDATE(
--        IF(POSITION1.start_dt >= transfer.start_dt,
--        POSITION1.start_dt,
--		  transfer.start_dt), 
--		  INTERVAL -1 MONTH
--      )
    ), 
    INTERVAL 1 DAY
  ) AS '시작일', 
  ADDDATE(
    LAST_DAY(
      ADDDATE(
        IF(
          transfer.End_dt = '9999-12-31', 
          CURDATE(), 
          IF(POSITION1.end_dt <= transfer.end_dt,
          POSITION1.end_dt,
			 transfer.End_dt)
        ), 
        INTERVAL -2 MONTH
      )
    ), 
    INTERVAL 1 DAY
  ) AS '종료일', 
  transfer.start_dt AS '관리시작일', 
  transfer.end_dt AS '관리종료일', 
  PERIOD_DIFF(
    DATE_FORMAT(
      IF(
        transfer.End_dt = '9999-12-31', 
        CURDATE(), 
        IF(POSITION1.end_dt <= transfer.end_dt,
        POSITION1.end_dt,
		  transfer.End_dt)
      ), 
      '%Y%m'
    ), 
    DATE_FORMAT(transfer.start_dt, '%Y%m')
  ) AS '기간', 
  title.title AS '직분', 
  position1.position AS '직책', 
  churchnm.church_gb AS '교회구분', 
  pstaff.lifeno AS '생명번호',
  transfer.church_sid AS '교회코드'
FROM 
  db_transfer transfer 
  LEFT JOIN db_pastoralstaff pstaff ON(pstaff.lifeno = transfer.lifeno) 
  LEFT JOIN db_churchlist_custom churchnm ON(
    REPLACE(transfer.church_sid, 'MC', 'MM') = churchnm.church_sid
  ) 
  LEFT JOIN db_title title ON(
    pstaff.lifeno = title.lifeno 
    AND LAST_DAY(transfer.start_dt) BETWEEN title.start_dt 
    AND title.end_dt
  ) 
  LEFT JOIN db_position position1 ON(
    pstaff.lifeno = position1.lifeno 
    AND LAST_DAY(transfer.start_dt) BETWEEN position1.start_dt 
    AND position1.end_dt
  ) 
WHERE 
  ((
    position1.position LIKE '%당%' 
    AND churchnm.church_gb IN ('MC', 'HBC', 'MM')
  )
  OR (
    position1.position LIKE '%지교회관리자%' 
    -- 2023.09.25 유럽2과 오류접수: HBC가 MC로 승격한 경우 선지자상세정보 안나오는 문제 해결
	 -- 이 경우 지관지 직책을 유지하면서 본교회에 있는 분들 데이터가 잘못 나올 가능성이 있지만 그것은 엄밀히 말하면 직책 데이터가 잘못 입력되어 있는 경우이므로 아래와 같이 수정하는 것이 좋다고 판단됨
    AND churchnm.church_gb IN ('MC', 'HBC')
  ))
  AND churchnm.ovs_dept <> ""
ORDER BY 
  생명번호, 
  관리종료일 DESC --  교회코드 DESC, ;

-- 뷰 op_system.v_pstaff_detail_flight 구조 내보내기
DROP VIEW IF EXISTS `v_pstaff_detail_flight`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v_pstaff_detail_flight`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v_pstaff_detail_flight` AS SELECT `flight`.`lifeno`        AS `생명번호`,
       `flight`.`flight_dt`     AS `방문일자`,
       `flight`.`visit_purpose` AS `방문목적`
FROM   `db_flight_schedule` `flight`
WHERE  `flight`.`destination` = '대한민국'
ORDER  BY `flight`.`lifeno`,
          `flight`.`flight_dt` ;

-- 뷰 op_system.v_pstaff_detail_title 구조 내보내기
DROP VIEW IF EXISTS `v_pstaff_detail_title`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v_pstaff_detail_title`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v_pstaff_detail_title` AS SELECT 
  title.lifeno AS '생명번호', 
  Ifnull(
    churchnm.church_nm, churchnm2.church_nm
  ) AS '교회명', 
  title.start_dt AS '임명일', 
  title.title AS '직분' 
FROM 
  db_title title 
  LEFT JOIN db_transfer transfer ON(
    title.lifeno = transfer.lifeno 
    AND title.start_dt BETWEEN transfer.start_dt 
    AND transfer.end_dt
  ) 
  LEFT JOIN db_churchlist churchnm ON(
    transfer.church_sid = churchnm.church_sid
  ) 
  LEFT JOIN db_pastoralwife wife ON(
    wife.lifeno = title.lifeno
  ) 
  LEFT JOIN db_transfer transfer2 ON(
    wife.lifeno_spouse = transfer2.lifeno 
    AND title.start_dt BETWEEN transfer2.start_dt 
    AND transfer2.end_dt
  ) 
  LEFT JOIN db_churchlist churchnm2 ON(
    transfer2.church_sid = churchnm2.church_sid
  ) 
ORDER BY 
  title.lifeno, 
  title.start_dt ;

-- 뷰 op_system.v_pstaff_detail_transfer 구조 내보내기
DROP VIEW IF EXISTS `v_pstaff_detail_transfer`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v_pstaff_detail_transfer`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v_pstaff_detail_transfer` AS -- ########################
-- 선지자 상세정보 발령이력
-- ########################


-- 본교회 전출입 이력
WITH MCTRANS AS
(
	SELECT 
	  transfer.lifeno AS '생명번호', 
	  transfer.start_dt AS '발령일', 
	  REPLACE(REPLACE(REPLACE(CONCAT(
	     IFNULL(title.title,''),
		  IF(title.title IS NULL,'', '/'),
		  IF(
		    position1.position LIKE '%관리자%' 
		    OR position1.position LIKE '%당%' 
		    OR position1.position LIKE '%동%', 
		    position1.position, 
		    theological.level
		  ))
		  ,'당회장대리','당대리'),'지교회관리자','지관자'),'예배소관리자','예관자') AS '직분/직책', 
	  churchnm.church_gb AS '교회구분',
	  REPLACE(
	  	IF(INSTR(REVERSE(churchnm.church_nm),' ') > 0,
		  RIGHT(churchnm.church_nm,INSTR(REVERSE(churchnm.church_nm),' ')-1),
		  churchnm.church_nm
		)
	  ,'_old','') AS '교회명',
	  churchnm.church_sid AS '교회코드',
	  IF(
	  	TIMESTAMPDIFF(MONTH, transfer.start_dt, LEAST(transfer.end_dt, POSITION1.end_dt, CURDATE())) >= 12,
	  	CONCAT(FLOOR(TIMESTAMPDIFF(MONTH, transfer.start_dt, LEAST(transfer.end_dt, POSITION1.end_dt, CURDATE()))/12),'년 ',
		  IF(MOD(TIMESTAMPDIFF(MONTH, transfer.start_dt, LEAST(transfer.end_dt, POSITION1.end_dt, CURDATE())),12)=0,'',
		  	CONCAT(MOD(TIMESTAMPDIFF(MONTH, transfer.start_dt, LEAST(transfer.end_dt, POSITION1.end_dt, CURDATE())),12),'개월')
		  )
		),
	  	CONCAT(TIMESTAMPDIFF(MONTH, transfer.start_dt, LEAST(transfer.end_dt, POSITION1.end_dt, CURDATE())),'개월')
	  ) '기간'
	  ,'MCTRANS' AS '이력구분'
	  ,LEAST(transfer.end_dt, IFNULL(POSITION1.end_dt,CURDATE())) AS '종료일'
	  ,title.title AS '직분'
	  ,position1.position AS '직책'
	  ,geo.country_nm_ko AS '국가명'
	FROM 
	  db_transfer transfer 
	  INNER JOIN db_pastoralstaff pstaff ON(
	    transfer.lifeno = pstaff.lifeno
	  ) 
	  LEFT JOIN db_title title ON(
	    pstaff.lifeno = title.lifeno 
	    AND Last_day(transfer.start_dt) BETWEEN title.start_dt 
	    AND title.end_dt
	  ) 
	  LEFT JOIN db_position position1 ON(
	    pstaff.lifeno = position1.lifeno 
	    AND Last_day(transfer.start_dt) BETWEEN position1.start_dt 
	    AND position1.end_dt
	  ) 
	  LEFT JOIN db_position2 position2 ON(
	    pstaff.lifeno - position2.lifeno 
	    AND Last_day(transfer.start_dt) BETWEEN position2.start_dt 
	    AND position2.end_dt
	  ) 
	  LEFT JOIN db_theological theological ON(
	    pstaff.lifeno = theological.lifeno 
	    AND Last_day(transfer.start_dt) BETWEEN theological.start_dt 
	    AND theological.end_dt
	  ) 
	  LEFT JOIN db_churchlist churchnm ON(
	    transfer.church_sid = churchnm.church_sid
	  )
	  LEFT JOIN db_geodata geo ON(
	  	 churchnm.geo_cd = geo.geo_cd
	  )
), BCTRANS AS (
	SELECT
	  b.lifeno AS '생명번호', 
	  b.start_dt AS '발령일', 
	  REPLACE(REPLACE(REPLACE(
	    CONCAT(
		 	IFNULL(title.title,''),
	 	   IF(title.title IS NULL,'', '/'),
			position1.position),
		 '당회장대리','당대리'),'지교회관리자','지관자'),'예배소관리자','예관자') AS '직분/직책', 
	  churchnm.church_gb AS '교회구분',
	  REPLACE(
	  	IF(INSTR(REVERSE(churchnm.church_nm),' ') > 0,
		  CONCAT(RIGHT(churchnm.church_nm,INSTR(REVERSE(churchnm.church_nm),' ')-1),'[',RIGHT(mainChurch.church_nm,INSTR(REVERSE(mainChurch.church_nm),' ')-1),']'),
		  churchnm.church_nm)
	  ,'_폐쇄','') AS '교회명',
	  churchnm.church_sid AS '교회코드',
	  REPLACE(REPLACE(REPLACE(CONCAT(
	  	TIMESTAMPDIFF(YEAR, b.start_dt, LEAST(b.end_dt, CURDATE())),'년',
		TIMESTAMPDIFF(MONTH, b.start_dt, LEAST(b.end_dt, CURDATE()))-TIMESTAMPDIFF(YEAR, b.start_dt, LEAST(b.end_dt, CURDATE()))*12,'개월'  
	  ),'0년',''),'년','년 '),' 0개월','') '기간'
	  ,'BCTRANS'
	  ,LEAST(b.End_dt, CURDATE())
	  ,title.title
	  ,position1.position
	  ,geo.country_nm_ko
	FROM 
		op_system.db_branchleader b
	  INNER JOIN db_pastoralstaff pstaff ON(
	    b.lifeno = pstaff.lifeno
	  ) 
	  LEFT JOIN db_title title ON(
	    pstaff.lifeno = title.lifeno 
	    AND Last_day(b.start_dt) BETWEEN title.start_dt 
	    AND title.end_dt
	  ) 
	  LEFT JOIN db_position position1 ON(
	    pstaff.lifeno = position1.lifeno 
	    AND LAST_DAY(b.start_dt) BETWEEN position1.start_dt 
	    AND position1.end_dt
	  ) 
	  LEFT JOIN db_churchlist churchnm ON(
	    b.church_sid = churchnm.church_sid
	  ) 
	  LEFT JOIN db_churchlist mainChurch ON(
	  	 churchnm.main_church_cd = mainChurch.church_sid
	  )
	  LEFT JOIN db_geodata geo ON(
	  	 churchnm.geo_cd = geo.geo_cd
	  )
	WHERE 
	  b.responsibility = '관리자'
), POSTRANS AS (
	SELECT 
	  position1.lifeno AS '생명번호', 
	  position1.start_dt AS '발령일', 
	  REPLACE(REPLACE(REPLACE(
	     CONCAT(
		  	  IFNULL(title.title,''),
	  	     IF(title.title IS NULL,'', '/'),
			  position1.position),
		  '당회장대리','당대리'),'지교회관리자','지관자'),'예배소관리자','예관자') AS '직분/직책', 
	  churchnm.church_gb AS '교회구분',
	  REPLACE(
	    IF(INSTR(REVERSE(churchnm.church_nm),' ')>0,
		 RIGHT(churchnm.church_nm,INSTR(REVERSE(churchnm.church_nm),' ')-1),
		 churchnm.church_nm)
	  ,'_old','') AS '교회명',
	  churchnm.church_sid AS '교회코드',
	  REPLACE(REPLACE(REPLACE(CONCAT(
	  	TIMESTAMPDIFF(YEAR, position1.start_dt, LEAST(position1.end_dt, transfer.end_dt, CURDATE())),'년',
		TIMESTAMPDIFF(MONTH, position1.start_dt, LEAST(position1.end_dt, transfer.end_dt, CURDATE()))-TIMESTAMPDIFF(YEAR, position1.start_dt, LEAST(position1.end_dt, transfer.end_dt, CURDATE()))*12,'개월'  
	  ),'0년',''),'년','년 '),' 0개월','') '기간'
	  ,'POSTRANS'
	  ,LEAST(position1.end_dt, transfer.end_dt, CURDATE())
	  ,title.title
	  ,position1.position
	  ,geo.country_nm_ko
	FROM 
	  db_position position1 
	  INNER JOIN db_pastoralstaff pstaff ON(
	    position1.lifeno = pstaff.lifeno
	  ) 
	  LEFT JOIN db_title title ON(
	    pstaff.lifeno = title.lifeno 
	    AND Last_day(position1.start_dt) BETWEEN title.start_dt 
	    AND title.end_dt
	  ) 
	  LEFT JOIN db_transfer transfer ON(
	    pstaff.lifeno = transfer.lifeno 
	    AND Last_day(position1.start_dt) BETWEEN transfer.start_dt 
	    AND transfer.end_dt
	  ) 
	  LEFT JOIN db_churchlist churchnm ON(
	    transfer.church_sid = churchnm.church_sid
	  ) 
	  LEFT JOIN db_geodata geo ON(
	  	 churchnm.geo_cd = geo.geo_cd
	  )
), TRANS AS (
	SELECT * FROM MCTRANS
	UNION
	SELECT * FROM POSTRANS
	WHERE `직분/직책` IS NOT NULL
		AND (`생명번호`, CONCAT(YEAR(`발령일`),'-',MONTH(`발령일`)), `교회코드`) NOT IN (SELECT `생명번호`, CONCAT(YEAR(`발령일`),'-',MONTH(`발령일`)), `교회코드` FROM MCTRANS)
		AND (`생명번호`, `발령일`, `직분/직책`) NOT IN (SELECT `생명번호`, `발령일`, `직분/직책` FROM BCTRANS)
--		AND NOT EXISTS (SELECT * FROM MCTRANS)
	UNION
	SELECT * FROM BCTRANS
	WHERE `직분/직책` IS NOT NULL
)

SELECT 
	TRANS.`생명번호`
	, TRANS.`발령일`
	, TRANS.`직분/직책`
	, TRANS.`교회구분`
	, CASE
	     WHEN TRANS.`국가명` = '대한민국'
	     THEN TRANS.`교회명`
	     WHEN TRANS.`국가명` <> geo.country_nm_ko
	     THEN CONCAT(TRANS.`국가명`, ' ' , TRANS.`교회명`)
	     ELSE TRANS.`교회명`
	  END AS '교회명'
	, TRANS.`교회코드`
	, TRANS.`기간`
	, TRANS.`이력구분`
	, TRANS.`종료일`
	, TRANS.`직분`
	, TRANS.`직책`
	, TRANS.`국가명` 
	, geo.country_nm_ko AS '현재국가' FROM TRANS
LEFT JOIN op_system.v0_pstaff_information_all pinfo
	ON TRANS.`생명번호` = pinfo.`생명번호`
LEFT JOIN op_system.db_churchlist churchlist
	ON pinfo.`교회코드` = churchlist.church_sid
LEFT JOIN op_system.db_geodata geo
	ON churchlist.geo_cd = geo.geo_cd
-- WHERE `직분/직책` IS NOT NULL
ORDER BY `생명번호`, `발령일` DESC, `종료일` 

-- 직책변동 이력



-- 지교회 전출입 이력

-- ORDER BY 
--  생명번호, 
--  발령일 ;

-- 뷰 op_system.v_search_titleposition 구조 내보내기
DROP VIEW IF EXISTS `v_search_titleposition`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v_search_titleposition`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v_search_titleposition` AS SELECT 
  pstaff.lifeno AS '생명번호', 
  IF(
    Locate(' ', churchlist.church_nm) > 0, 
    Reverse(
      LEFT(
        Reverse(churchlist.church_nm), 
        Locate(
          ' ', 
          Reverse(churchlist.church_nm)
        ) -1
      )
    ), 
    churchlist.church_nm
  ) AS '교회명', 
  church_admin.church_nm_en AS 영문교회명, 
  IF(
    branchnm.church_nm IS NULL, 
    IF(
      Locate(' ', churchlist.church_nm) > 0, 
      Reverse(
        LEFT(
          Reverse(churchlist.church_nm), 
          Locate(
            ' ', 
            Reverse(churchlist.church_nm)
          ) -1
        )
      ), 
      churchlist.church_nm
    ), 
    branchnm.church_nm
  ) AS '지교회명', 
  IF(
    branch_admin.church_nm_en IS NULL, 
    church_admin.church_nm_en, branch_admin.church_nm_en
  ) AS '영문지교회명', 
-- 선교국가 검색 시 본교회 기준으로 검색할 일이 더 많으므로 
-- 기존 지교회 기반 검색에서 본교회 기반 검색으로 수정함(2023-01-11)
  IFNULL(geoMainOfBranch.country_nm_ko, geo.country_nm_ko) AS '선교국가', 
  Concat(
    pstaff.name_ko, 
    Ifnull(
      Concat(
        '(', 
        LEFT(title.title, 1), 
        ')'
      ), 
      ''
    )
  ) AS '한글이름(직분)', 
  pstaff.name_en AS '영문이름', 
  Ifnull(
    IF(
      position.position LIKE '%관리자%' 
      OR position.position LIKE '%당%' 
      OR position.position LIKE '%동%', 
      position.position, 
      theological.level
    ), 
    IF(
      position2.position2 IS NOT NULL, 
      IF(
        position.position IS NULL, '직책없음', 
        position.position
      ), 
      '직책없음'
    )
  ) AS '직책', 
  position2.position2 AS '직책2', 
  pstaff.birthday AS '생년월일', 
  pstaff.nationality AS '국적', 
  Ifnull(
    IF(
      pstaff.appo_ovs IS NULL, appoint.start_dt, 
      pstaff.appo_ovs
    ), 
    '없음'
  ) AS '(해외)최초발령일', 
  Ifnull(
    IF(
      position.position IN (
        '당회장', '당회장대리', '동역'
      ), 
      CASE WHEN position.position = '동역' THEN IF(
        position.start_dt >= belong.start_dt, 
        position.start_dt, belong.start_dt
      ) ELSE belong.start_dt END, 
      NULL
    ), 
    '없음'
  ) AS '현당회발령일', 
  IF(
    Curdate() < pstaff.wedding_dt, 
    NULL, 
    spouse.lifeno
  ) AS '배우자생번', 
  IF(
    Curdate() < pstaff.wedding_dt, 
    '', 
    Concat(
      spouse.name_ko, 
      Ifnull(
        Concat(
          '(', 
          LEFT(title_spouse.title, 1), 
          ')'
        ), 
        ''
      )
    )
  ) AS '사모한글이름(직분)', 
  IF(
    Curdate() < pstaff.wedding_dt, 
    '', 
    spouse.name_en
  ) AS '사모영문이름', 
  IF(
    Curdate() < pstaff.wedding_dt, 
    '', 
    -- OR pstaff.wedding_dt IS NULL
    IF(
      spouse.lifeno IS NOT NULL, spouseposition.position_spouse, 
      NULL
    )
  ) AS '사모직책', 
  IF(
    Curdate() < pstaff.wedding_dt, 
    '', 
    IF(
      spouse.lifeno IS NOT NULL, position2_spouse.position2, 
      NULL
    )
  ) AS '사모직책2', 
  IF(
    Curdate() < pstaff.wedding_dt, 
    NULL, 
    spouse.birthday
  ) AS '배우자 생년월일', 
  IF(
    Ifnull(
      position.position, theological.level
    ) LIKE '당%' 
    OR Ifnull(
      position.position, theological.level
    ) LIKE '동%', 
    NULL, 
    theological.level
  ) AS '생도기수', 
  title.title AS '직분', 
  IF(
    Last_day(
      Curdate() + INTERVAL -1 month
    ) < pstaff.wedding_dt 
    OR pstaff.wedding_dt IS NULL, 
    NULL, 
    title_spouse.title
  ) AS '사모직분', 
  pstaff.baptism AS '침례권', 
  unionnm.union_nm AS '연합회', 
  atten.once_all AS '전체1회', 
  atten.once_stu AS '학생1회', 
  IF(churchlist.church_gb = 'HBC', atten.once_all, attenbranch.once_all) AS '지교회전체1회', -- HBC를 지교회처럼 취급하기 위한 조치
  IF(churchlist.church_gb = 'HBC', atten.once_stu, attenbranch.once_stu) AS '지교회학생1회', 
  cntbc.관리지교회 AS '관리지교회', 
  cntbc.관리예배소 AS '관리예배소', 
  cntleader.동역 AS '동역', 
  cntleader.지교회관리자 AS '지교회관리자', 
  cntleader.예배소관리자 AS '예배소관리자', 
  cntleader.예비생도 AS '예비생도', 
  spouse.nationality AS '사모국적', 
  position2.start_dt AS '직책2시작일', 
  position2_spouse.start_dt AS '사모직책2시작일', 
  atten2.once_all AS '전체1회(2달 전)', 
  atten2.once_stu AS '학생1회(2달 전)', 
  IF(churchlist.church_gb = 'HBC', atten2.once_all, attenbranch2.once_all) AS '지교회전체1회(2달 전)', -- HBC를 지교회처럼 취급하기 위한 조치
  IF(churchlist.church_gb = 'HBC', atten2.once_stu, attenbranch2.once_stu) AS '지교회학생1회(2달 전)', -- HBC를 지교회처럼 취급하기 위한 조치
  pstaff.ovs_dept AS '관리부서', 
  unionnm.sort_order AS '연합회 정렬순서', 
  churchlist.sort_order AS '본교회 정렬순서', 
  IFNULL(
	  REPLACE(
	    REPLACE(
	      branchnm.church_gb, 'PBC', '예배소'
	    ), 
	    'BC', 
	    '지교회'
	  ), 
	  REPLACE(REPLACE(churchlist.church_gb,'MC',''),'HBC','지교회') -- HBC를 지교회처럼 취급하기 위한 조치
  ) AS '교회구분', 
  branchnm.sort_order AS '지교회 정렬순서', 
  churchlist.church_nm AS '교회명(전체)',
  IFNULL(
  	  branchlist.responsibility, IF(churchlist.church_gb = 'HBC', '관리자', NULL) -- HBC를 지교회처럼 취급하기 위한 조치
  ) AS '지교회역할' 
FROM 
  op_system.db_pastoralstaff pstaff 
  LEFT JOIN op_system.db_title title ON(
    pstaff.lifeno = title.lifeno 
    AND Curdate() BETWEEN title.start_dt 
    AND title.end_dt
  ) 
  LEFT JOIN op_system.db_position position ON(
    pstaff.lifeno = position.lifeno 
    AND Curdate() BETWEEN position.start_dt 
    AND position.end_dt
  ) 
  LEFT JOIN op_system.db_position2 position2 ON(
    pstaff.lifeno = position2.lifeno 
    AND Curdate() BETWEEN position2.start_dt 
    AND position2.end_dt
  ) 
  LEFT JOIN op_system.db_transfer belong ON(
    pstaff.lifeno = belong.lifeno 
    AND Curdate() BETWEEN belong.start_dt 
    AND belong.end_dt
  ) 
  LEFT JOIN op_system.db_churchlist churchlist ON(
    belong.church_sid = churchlist.church_sid
  ) 
  LEFT JOIN op_system.db_history_church_establish churchesta ON(
    churchlist.church_sid = churchesta.church_sid
  ) 
  LEFT JOIN op_system.db_union uniondb ON(
    churchesta.church_sid_custom = uniondb.church_sid_custom
    AND CURDATE() BETWEEN uniondb.start_dt AND uniondb.end_dt
  ) 
  LEFT JOIN op_system.a_union unionnm ON(
  	 uniondb.union = unionnm.union_cd
  ) 
  LEFT JOIN op_system.db_attendance atten ON(
    churchlist.church_sid = atten.church_sid 
    AND atten.attendance_dt = Last_day(
      Curdate() + INTERVAL -2 month
    ) + INTERVAL 1 day
  ) 
  LEFT JOIN op_system.db_attendance atten2 ON(
    churchlist.church_sid = atten2.church_sid 
    AND atten2.attendance_dt = Last_day(
      Curdate() + INTERVAL -3 month
    ) + INTERVAL 1 day
  ) 
  LEFT JOIN op_system.db_branchleader branchlist ON(
    pstaff.lifeno = branchlist.lifeno 
    AND Curdate() BETWEEN branchlist.start_dt 
    AND branchlist.end_dt
  ) 
  LEFT JOIN op_system.db_churchlist branchnm ON(
    branchlist.church_sid = branchnm.church_sid
  ) 
  LEFT JOIN op_system.db_churchlist mainchurchnm ON(
  	 branchnm.main_church_cd = mainchurchnm.church_sid
  )
  LEFT JOIN op_system.db_geodata geo ON geo.geo_cd = churchlist.geo_cd 
  LEFT JOIN op_system.db_geodata geoBranch ON geoBranch.geo_cd = branchnm.geo_cd 
  LEFT JOIN op_system.db_geodata geoMainOfBranch ON geoMainOfBranch.geo_cd = mainchurchnm.geo_cd
  LEFT JOIN op_system.db_attendance attenbranch ON(
    branchlist.church_sid = attenbranch.church_sid 
    AND attenbranch.attendance_dt = Last_day(
      Curdate() + INTERVAL -2 month
    ) + INTERVAL 1 day
  ) 
  LEFT JOIN op_system.db_attendance attenbranch2 ON(
    branchlist.church_sid = attenbranch2.church_sid 
    AND attenbranch2.attendance_dt = Last_day(
      Curdate() + INTERVAL -3 month
    ) + INTERVAL 1 day
  ) 
  LEFT JOIN op_system.db_theological theological ON(
    pstaff.lifeno = theological.lifeno 
    AND Curdate() BETWEEN theological.start_dt 
    AND theological.end_dt
  ) 
  LEFT JOIN op_system.db_pastoralwife spouse ON(
    pstaff.lifeno = spouse.lifeno_spouse
  ) 
  LEFT JOIN op_system.db_position2 position2_spouse ON(
    spouse.lifeno = position2_spouse.lifeno 
    AND Curdate() BETWEEN position2_spouse.start_dt 
    AND position2_spouse.end_dt
  ) 
  LEFT JOIN op_system.a_position_spouse spouseposition ON(
--    Ifnull(
--      position.position, theological.level
--    ) = spouseposition.position
		Ifnull(
		    IF(
		      position.position LIKE '%관리자%' 
		      OR position.position LIKE '%당%' 
		      OR position.position LIKE '%동%', 
		      position.position, 
		      theological.level
		    ), 
		    IF(
		      position2.position2 IS NOT NULL, 
		      IF(
		        position.position IS NULL, '직책없음', 
		        position.position
		      ), 
		      '직책없음'
		    )
		  ) = spouseposition.position
  ) 
  LEFT JOIN (
    SELECT 
      op_system.db_position.lifeno AS lifeno, 
      Min(op_system.db_position.start_dt) AS start_dt, 
      op_system.db_position.position AS Position 
    FROM 
      op_system.db_position 
    WHERE 
      op_system.db_position.position = '당회장' 
      OR op_system.db_position.position = '당회장대리' 
      OR op_system.db_position.position = '동역' 
    GROUP BY 
      op_system.db_position.lifeno
  ) appoint ON(appoint.lifeno = pstaff.lifeno) 
  LEFT JOIN (
    SELECT 
      op_system.db_title.title_cd AS title_cd, 
      op_system.db_title.lifeno AS LifeNo, 
      op_system.db_title.start_dt AS Start_dt, 
      op_system.db_title.end_dt AS End_dt, 
      op_system.db_title.title AS Title 
    FROM 
      op_system.db_title
  ) title_spouse ON(
    title_spouse.lifeno = spouse.lifeno 
    AND Curdate() BETWEEN title_spouse.start_dt 
    AND title_spouse.end_dt
  ) 
  LEFT JOIN op_system.a_churchlist_admin church_admin ON(
    churchlist.church_sid = church_admin.church_sid
  ) 
  LEFT JOIN op_system.a_churchlist_admin branch_admin ON(
    branchlist.church_sid = branch_admin.church_sid
  ) 
  LEFT JOIN (
    SELECT 
      a.main_church AS main_church, 
      Count(
        IF(
          a.church_gb NOT LIKE '%LBC%', NULL, 
          a.church_gb
        )
      ) AS 관리지교회, 
      Count(
        IF(
          a.church_gb NOT LIKE '%LPBC%', NULL, 
          a.church_gb
        )
      ) AS 관리예배소 
    FROM 
      op_system.a_churchlist_admin a 
    GROUP BY 
      a.main_church
  ) cntbc ON(
    churchlist.church_nm = cntbc.main_church
  ) 
  LEFT JOIN (
    SELECT 
      a.교회명 AS 교회명, 
      Count(
        IF(
          a.직책 NOT LIKE '%동역%', NULL, 
          a.직책
        )
      ) AS 동역, 
      Count(
        IF(
          a.직책 NOT LIKE '%지교회%', NULL, 
          a.직책
        )
      ) AS 지교회관리자, 
      Count(
        IF(
          a.직책 NOT LIKE '%예배소%', NULL, 
          a.직책
        )
      ) AS 예배소관리자, 
      Count(
        IF(
          a.직책 NOT LIKE '%생도%', NULL, 
          a.직책
        )
      ) AS 예비생도 
    FROM 
      op_system.v0_pstaff_information a 
    GROUP BY 
      a.교회명
  ) cntleader ON(
    churchlist.church_nm = cntleader.교회명
  ) 
WHERE 
  churchlist.church_nm IS NOT NULL 
--  AND church_admin.country IS NOT NULL 
  AND geo.country_nm_ko IS NOT NULL
  AND (
    IF(
      position.position LIKE '%관리자%' 
      OR position.position LIKE '%당%' 
      OR position.position LIKE '%동%', 
      position.position, 
      theological.level
    ) IS NOT NULL 
    OR position2.position2 IS NOT NULL
  ) 
  AND (
    position2.position2 IS NOT NULL 
    OR IF(
      position.position LIKE '%관리자%' 
      OR position.position LIKE '%당%' 
      OR position.position LIKE '%동%', 
      position.position, 
      theological.level
    ) NOT LIKE '%역장%'
  ) 
ORDER BY 
  IF(
    Locate(' ', churchlist.church_nm) > 0, 
    Reverse(
      LEFT(
        Reverse(churchlist.church_nm), 
        Locate(
          ' ', 
          Reverse(churchlist.church_nm)
        ) -1
      )
    ), 
    churchlist.church_nm
  ), 
  Ifnull(
    IF(
      pstaff.appo_ovs IS NULL, appoint.start_dt, 
      pstaff.appo_ovs
    ), 
    '없음'
  ) ;

-- 뷰 op_system.v_transfer_history 구조 내보내기
DROP VIEW IF EXISTS `v_transfer_history`;
-- 임시 테이블을 제거하고 최종 VIEW 구조를 생성
DROP TABLE IF EXISTS `v_transfer_history`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `v_transfer_history` AS SELECT 
  transfer.lifeno AS '생명번호', 
  vinfo.`교회명` AS '교회명', 
  vinfo.`한글이름(직분)` AS '선지자이름(직분)', 
  vinfo.`직책` AS '선지자직책', 
  pinfo.birthday AS '선지자생년월일', 
  pinfo.nationality AS '선지자국적', 
  pinfo.home AS '선지자고향', 
  pinfo.family AS '선지자가족', 
  pinfo.health AS '선지자건강', 
  pinfo.other AS '선지자기타', 
  flight1.flight_dt AS '선지자마지막방문일', 
  flight1.visit_purpose AS '선지자방문목적', 
  winfo.lifeno AS '사모생번', 
  vinfo.`사모한글이름(직분)` AS '사모이름(직분)', 
  vinfo.`사모직책` AS '사모직책', 
  winfo.birthday AS '사모생년월일', 
  winfo.nationality AS '사모국적', 
  winfo.home AS '사모고향', 
  winfo.family AS '사모가족', 
  winfo.health AS '사모건강', 
  winfo.other AS '사모기타', 
  flight2.flight_dt AS '사모마지막방문일', 
  flight2.visit_purpose AS '사모방문목적', 
  pinfo.wedding_dt AS '혼인일', 
  vinfo.`(해외)최초발령일` AS '(해외)최초발령일', 
  vinfo.`현당회발령일` AS '현당회발령일', 
  transfer.start_dt AS '발령일', 
  positionfrom.position AS '전출직책', 
  LEFT(titlefrom.title, 1) AS '전출직분', 
  churchfrom.church_nm AS '전출교회', 
  positionto.position AS '전입직책', 
  LEFT(titleto.title, 1) AS '전입직분', 
  churchto.church_nm AS '전입교회',
  pinfo.lifeno_child1 AS '자녀1 생명번호',
  pinfo.birthday_child1 AS '자녀1 생년월일',
  pinfo.lifeno_child2 AS '자녀2 생명번호',
  pinfo.birthday_child2 AS '자녀2 생년월일',
  pinfo.lifeno_child3 AS '자녀3 생명번호',
  pinfo.birthday_child3 AS '자녀3 생년월일'
FROM 
  op_system.db_transfer transfer 
  LEFT JOIN op_system.db_transfer transferfrom ON(
    transfer.lifeno = transferfrom.lifeno 
    AND transfer.start_dt + INTERVAL -1 day BETWEEN transferfrom.start_dt 
    AND transferfrom.end_dt
  ) 
  LEFT JOIN op_system.db_title titleto ON(
    transfer.lifeno = titleto.lifeno 
    AND transfer.start_dt BETWEEN titleto.start_dt 
    AND titleto.end_dt
  ) 
  LEFT JOIN op_system.db_title titlefrom ON(
    transfer.lifeno = titlefrom.lifeno 
    AND transfer.start_dt + INTERVAL -1 day BETWEEN titlefrom.start_dt 
    AND titlefrom.end_dt
  ) 
  LEFT JOIN op_system.db_position positionto ON(
    transfer.lifeno = positionto.lifeno 
    AND transfer.start_dt BETWEEN positionto.start_dt 
    AND positionto.end_dt
  ) 
  LEFT JOIN op_system.db_position positionfrom ON(
    transfer.lifeno = positionfrom.lifeno 
    AND transfer.start_dt + INTERVAL -1 day BETWEEN positionfrom.start_dt 
    AND positionfrom.end_dt
  ) 
  LEFT JOIN op_system.db_churchlist churchto ON(
    transfer.church_sid = churchto.church_sid
  ) 
  LEFT JOIN op_system.db_churchlist churchfrom ON(
    transferfrom.church_sid = churchfrom.church_sid
  ) 
  LEFT JOIN op_system.db_pastoralstaff pinfo ON(
    transfer.lifeno = pinfo.lifeno
  ) 
  LEFT JOIN op_system.db_pastoralwife winfo ON(
    pinfo.lifeno = winfo.lifeno_spouse
  ) 
  LEFT JOIN op_system.v0_pstaff_information vinfo ON(
    transfer.lifeno = vinfo.생명번호
  ) 
  LEFT JOIN (
    SELECT 
      f_sche1.lifeno AS lifeno, 
      f_sche1.flight_dt AS flight_dt, 
      f_sche1.departure AS departure, 
      f_sche1.destination AS destination, 
      f_sche1.visit_purpose AS visit_purpose 
    FROM 
      op_system.db_flight_schedule f_sche1 
    WHERE 
      (
        f_sche1.lifeno, f_sche1.flight_dt
      ) IN (
        SELECT 
          a.lifeno, 
          Max(a.flight_dt) 
        FROM 
          op_system.db_flight_schedule a 
        WHERE 
          a.destination = '대한민국' 
        GROUP BY 
          a.lifeno
      )
  ) flight1 ON(
    transfer.lifeno = flight1.lifeno
  ) 
  LEFT JOIN (
    SELECT 
      f_sche2.lifeno AS lifeno, 
      f_sche2.flight_dt AS flight_dt, 
      f_sche2.departure AS departure, 
      f_sche2.destination AS destination, 
      f_sche2.visit_purpose AS visit_purpose 
    FROM 
      op_system.db_flight_schedule f_sche2 
    WHERE 
      (
        f_sche2.lifeno, f_sche2.flight_dt
      ) IN (
        SELECT 
          a.lifeno, 
          Max(a.flight_dt) 
        FROM 
          op_system.db_flight_schedule a 
        WHERE 
          a.destination = '대한민국' 
        GROUP BY 
          a.lifeno
      )
  ) flight2 ON(
    winfo.lifeno = flight2.lifeno
  ) 
WHERE 
  transferfrom.church_sid IS NOT NULL 
  AND vinfo.`한글이름(직분)` IS NOT NULL 
UNION 
SELECT 
  position.lifeno AS '생명번호', 
  vinfo.`교회명` AS '교회명', 
  vinfo.`한글이름(직분)` AS '선지자이름(직분)', 
  vinfo.`직책` AS '선지자직책', 
  pinfo.birthday AS '선지자생년월일', 
  pinfo.nationality AS '선지자국적', 
  pinfo.home AS '선지자고향', 
  pinfo.family AS '선지자가족', 
  pinfo.health AS '선지자건강', 
  pinfo.other AS '선지자기타', 
  flight1.flight_dt AS '선지자마지막방문일', 
  flight1.visit_purpose AS '선지자방문목적', 
  winfo.lifeno AS '사모생번', 
  vinfo.`사모한글이름(직분)` AS '사모이름(직분)', 
  vinfo.`사모직책` AS '사모직책', 
  winfo.birthday AS '사모생년월일', 
  winfo.nationality AS '사모국적', 
  winfo.home AS '사모고향', 
  winfo.family AS '사모가족', 
  winfo.health AS '사모건강', 
  winfo.other AS '사모기타', 
  flight2.flight_dt AS '사모마지막방문일', 
  flight2.visit_purpose AS '사모방문목적', 
  pinfo.wedding_dt AS '혼인일', 
  vinfo.`(해외)최초발령일` AS '(해외)최초발령일', 
  vinfo.`현당회발령일` AS '현당회발령일', 
  position.start_dt AS '발령일', 
  positionfrom.position AS '전출직책', 
  LEFT(titlefrom.title, 1) AS '전출직분', 
  churchfrom.church_nm AS '전출교회', 
  position.position AS '전입직책', 
  LEFT(titleto.title, 1) AS '전입직분', 
  churchto.church_nm AS '전입교회',
  pinfo.lifeno_child1 AS '자녀1 생명번호',
  pinfo.birthday_child1 AS '자녀1 생년월일',
  pinfo.lifeno_child2 AS '자녀2 생명번호',
  pinfo.birthday_child2 AS '자녀2 생년월일',
  pinfo.lifeno_child3 AS '자녀3 생명번호',
  pinfo.birthday_child3 AS '자녀3 생년월일'
FROM 
  op_system.db_position position 
  LEFT JOIN op_system.db_position positionfrom ON(
    position.lifeno = positionfrom.lifeno 
    AND position.start_dt + INTERVAL -1 day BETWEEN positionfrom.start_dt 
    AND positionfrom.end_dt
  ) 
  LEFT JOIN op_system.db_title titleto ON(
    position.lifeno = titleto.lifeno 
    AND position.start_dt BETWEEN titleto.start_dt 
    AND titleto.end_dt
  ) 
  LEFT JOIN op_system.db_title titlefrom ON(
    position.lifeno = titlefrom.lifeno 
    AND position.start_dt + INTERVAL -1 day BETWEEN titlefrom.start_dt 
    AND titlefrom.end_dt
  ) 
  LEFT JOIN op_system.db_transfer transferto ON(
    position.lifeno = transferto.lifeno 
    AND position.start_dt BETWEEN transferto.start_dt 
    AND transferto.end_dt
  ) 
  LEFT JOIN op_system.db_transfer transferfrom ON(
    position.lifeno = transferfrom.lifeno 
    AND position.start_dt + INTERVAL -1 day BETWEEN transferfrom.start_dt 
    AND transferfrom.end_dt
  ) 
  LEFT JOIN op_system.db_churchlist churchto ON(
    transferto.church_sid = churchto.church_sid
  ) 
  LEFT JOIN op_system.db_churchlist churchfrom ON(
    transferfrom.church_sid = churchfrom.church_sid
  ) 
  LEFT JOIN op_system.db_pastoralstaff pinfo ON(
    position.lifeno = pinfo.lifeno
  ) 
  LEFT JOIN op_system.db_pastoralwife winfo ON(
    pinfo.lifeno = winfo.lifeno_spouse
  ) 
  LEFT JOIN op_system.v0_pstaff_information vinfo ON(
    position.lifeno = vinfo.생명번호
  ) 
  LEFT JOIN (
    SELECT 
      f_sche1.lifeno AS lifeno, 
      f_sche1.flight_dt AS flight_dt, 
      f_sche1.departure AS departure, 
      f_sche1.destination AS destination, 
      f_sche1.visit_purpose AS visit_purpose 
    FROM 
      op_system.db_flight_schedule f_sche1 
    WHERE 
      (
        f_sche1.lifeno, f_sche1.flight_dt
      ) IN (
        SELECT 
          a.lifeno, 
          Max(a.flight_dt) 
        FROM 
          op_system.db_flight_schedule a 
        WHERE 
          a.destination = '대한민국' 
        GROUP BY 
          a.lifeno
      )
  ) flight1 ON(
    position.lifeno = flight1.lifeno
  ) 
  LEFT JOIN (
    SELECT 
      f_sche2.lifeno AS lifeno, 
      f_sche2.flight_dt AS flight_dt, 
      f_sche2.departure AS departure, 
      f_sche2.destination AS destination, 
      f_sche2.visit_purpose AS visit_purpose 
    FROM 
      op_system.db_flight_schedule f_sche2 
    WHERE 
      (
        f_sche2.lifeno, f_sche2.flight_dt
      ) IN (
        SELECT 
          a.lifeno, 
          Max(a.flight_dt) 
        FROM 
          op_system.db_flight_schedule a 
        WHERE 
          a.destination = '대한민국' 
        GROUP BY 
          a.lifeno
      )
  ) flight2 ON(
    winfo.lifeno = flight2.lifeno
  ) 
WHERE 
  vinfo.`한글이름(직분)` IS NOT NULL 
  AND (
    position.position LIKE '%당%' 
    OR position.position LIKE '%동%' 
    OR position.position LIKE '%관리자%'
  ) 
UNION 
SELECT 
  transfer2.lifeno AS '생명번호', 
  vinfo2.교회명 AS '교회명', 
  vinfo2.`한글이름(직분)` AS '선지자이름(직분)', 
  vinfo2.직책 AS '선지자직책', 
  pinfo2.birthday AS '선지자생년월일', 
  pinfo2.nationality AS '선지자국적', 
  pinfo2.home AS '선지자고향', 
  pinfo2.family AS '선지자가족', 
  pinfo2.health AS '선지자건강', 
  pinfo2.other AS '선지자기타', 
  flight1.flight_dt AS '선지자마지막방문일', 
  flight1.visit_purpose AS '선지자방문목적', 
  winfo2.lifeno AS '사모생번', 
  vinfo2.`사모한글이름(직분)` AS '사모이름(직분)', 
  vinfo2.`사모직책` AS '사모직책', 
  winfo2.birthday AS '사모생년월일', 
  winfo2.nationality AS '사모국적', 
  winfo2.home AS '사모고향', 
  winfo2.family AS '사모가족', 
  winfo2.health AS '사모건강', 
  winfo2.other AS '사모기타', 
  flight2.flight_dt AS '사모마지막방문일', 
  flight2.visit_purpose AS '사모방문목적', 
  pinfo2.wedding_dt AS '혼인일', 
  vinfo2.`(해외)최초발령일` AS '(해외)최초발령일', 
  vinfo2.`현당회발령일` AS '현당회발령일', 
  transfer2.start_dt AS '발령일', 
  NULL AS '전출직책', 
  NULL AS '전출직분', 
  Concat(transfer2.title, ' 임명') AS '전출교회', 
  NULL AS '전입직책', 
  NULL AS '전입직분', 
  NULL AS '전입교회',
  pinfo2.lifeno_child1 AS '자녀1 생명번호',
  pinfo2.birthday_child1 AS '자녀1 생년월일',
  pinfo2.lifeno_child2 AS '자녀2 생명번호',
  pinfo2.birthday_child2 AS '자녀2 생년월일',
  pinfo2.lifeno_child3 AS '자녀3 생명번호',
  pinfo2.birthday_child3 AS '자녀3 생년월일'
FROM 
  op_system.db_title transfer2 
  LEFT JOIN op_system.db_pastoralstaff pinfo2 ON(
    transfer2.lifeno = pinfo2.lifeno
  ) 
  LEFT JOIN op_system.db_pastoralwife winfo2 ON(
    pinfo2.lifeno = winfo2.lifeno_spouse
  ) 
  LEFT JOIN op_system.v0_pstaff_information vinfo2 ON(
    transfer2.lifeno = vinfo2.생명번호
  ) 
  LEFT JOIN (
    SELECT 
      f_sche1.lifeno AS lifeno, 
      f_sche1.flight_dt AS flight_dt, 
      f_sche1.departure AS departure, 
      f_sche1.destination AS destination, 
      f_sche1.visit_purpose AS visit_purpose 
    FROM 
      op_system.db_flight_schedule f_sche1 
    WHERE 
      (
        f_sche1.lifeno, f_sche1.flight_dt
      ) IN (
        SELECT 
          a.lifeno, 
          Max(a.flight_dt) 
        FROM 
          op_system.db_flight_schedule a 
        WHERE 
          a.destination = '대한민국' 
        GROUP BY 
          a.lifeno
      )
  ) flight1 ON(
    transfer2.lifeno = flight1.lifeno
  ) 
  LEFT JOIN (
    SELECT 
      f_sche2.lifeno AS lifeno, 
      f_sche2.flight_dt AS flight_dt, 
      f_sche2.departure AS departure, 
      f_sche2.destination AS destination, 
      f_sche2.visit_purpose AS visit_purpose 
    FROM 
      op_system.db_flight_schedule f_sche2 
    WHERE 
      (
        f_sche2.lifeno, f_sche2.flight_dt
      ) IN (
        SELECT 
          a.lifeno, 
          Max(a.flight_dt) 
        FROM 
          op_system.db_flight_schedule a 
        WHERE 
          a.destination = '대한민국' 
        GROUP BY 
          a.lifeno
      )
  ) flight2 ON(
    winfo2.lifeno = flight2.lifeno
  ) 
WHERE 
  vinfo2.`한글이름(직분)` IS NOT NULL 
ORDER BY 
  생명번호, 
  발령일 ;

/*!40101 SET SQL_MODE=IFNULL(@OLD_SQL_MODE, '') */;
/*!40014 SET FOREIGN_KEY_CHECKS=IF(@OLD_FOREIGN_KEY_CHECKS IS NULL, 1, @OLD_FOREIGN_KEY_CHECKS) */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
