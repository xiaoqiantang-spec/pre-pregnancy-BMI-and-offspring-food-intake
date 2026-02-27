###############################################################
## 个人食物重新匹配（CHNS foodcode -> FCT1991/2002/2004）
## ✅ 整合后一键运行版（含“食物大类 food_category”）
##
## 输入：
## - data1: 个体-食物条目（纵向），含 wave / vd / foodcode（及其他列无妨）
## - FCT 汇总表：1991(sheet1) / 2002(sheet2) / 2004(sheet3)
##
## 规则：
## - wave < 2004 : 用 FCT1991
## - wave >= 2004: FCT2002 优先 + FCT2004 兜底（逐字段 coalesce）
## - FCT 表中你手动填的 "NA" 视为缺失：统一转成真正 NA
## - 只合并“三张 FCT 都有的营养素变量”（列名交集）
## - 额外加入“食物大类 food_category”（在各自 FCT 内生成，不受交集限制）
##
## 输出对象：
## - data_food_matched
## - qc_match_rate_by_wave
## - unmatched_foodcodes
##
## 可选导出 Excel：取消 write_xlsx 注释
###############################################################

rm(list = ls())

suppressPackageStartupMessages({
  library(dplyr)
  library(stringr)
  library(readxl)
  library(tidyr)
  library(writexl)
})

###############################################################
## 0) 路径 + 读入个人食物条目
###############################################################
data1 <- read_excel(
  "E:/博士课题/CHNS/CHNS_2013/初次整理/最终使用数据/食物变量再次处理/个人食物量表全匹配/匹配子代食物code.xlsx",
  sheet = 1
)
stopifnot(exists("data1"))

path_fct <- "E:/博士课题/CHNS/CHNS_2013/食物成分表/整理好的食物成分表/食物成分表汇总.xls"

###############################################################
## 1) 工具函数
###############################################################

## 1.1 把各种伪缺失字符统一转 NA（含你手动填的 "NA"）
clean_na_char <- function(x) {
  if (!is.character(x)) x <- as.character(x)
  x <- trimws(x)
  x[x %in% c("", " ", ".", "NA","Na","na","N/A","n/a","NULL","null","None","none")] <- NA
  x
}

## 1.2 foodcode 统一函数（避免 11101 / 11101.0 / " 11101 "）
normalize_foodcode <- function(x) {
  x <- clean_na_char(x)
  x <- sub("\\.0+$", "", x)       # 去掉 .0 .00
  x <- gsub("[^0-9]", "", x)      # 只保留数字
  x <- sub("^0+", "", x)          # 去掉前导0
  x[x == ""] <- NA
  x
}

## 1.3 食物大类映射：1991（A01~A28）
map_cat_1991 <- function(bigcode) {
  dplyr::case_when(
    bigcode == "A01" ~ "谷类及制品",
    bigcode == "A02" ~ "干豆类及制品",
    bigcode == "A03" ~ "鲜豆类",
    bigcode == "A04" ~ "根茎类及制品",
    bigcode == "A05" ~ "嫩茎、叶、花类",
    bigcode == "A06" ~ "瓜类",
    bigcode == "A07" ~ "茄果类",
    bigcode == "A08" ~ "咸菜类",
    bigcode == "A09" ~ "菌藻类",
    bigcode == "A10" ~ "鲜果及干果类",
    bigcode == "A11" ~ "坚果类",
    bigcode == "A12" ~ "畜肉类及制品",
    bigcode == "A13" ~ "禽肉类及制品",
    bigcode == "A14" ~ "乳类及制品",
    bigcode == "A15" ~ "婴儿配方食品及辅助食品",
    bigcode == "A16" ~ "蛋类及制品",
    bigcode == "A17" ~ "鱼类",
    bigcode == "A18" ~ "软体动物类",
    bigcode == "A19" ~ "虾蟹类",
    bigcode == "A20" ~ "油脂类",
    bigcode == "A21" ~ "糕点及小吃类",
    bigcode == "A22" ~ "茶及饮料",
    bigcode == "A23" ~ "酒类",
    bigcode == "A24" ~ "糖及制品",
    bigcode == "A25" ~ "淀粉类及制品",
    bigcode == "A26" ~ "调味品类",
    bigcode == "A27" ~ "药用食物类",
    bigcode == "A28" ~ "杂类",
    TRUE ~ NA_character_
  )
}

## 1.4 食物大类映射：2002/2004（01-~21-）
map_cat_2002 <- function(prefix3) {
  dplyr::case_when(
    prefix3 == "01-" ~ "谷类及制品",
    prefix3 == "02-" ~ "薯类及淀粉及制品",
    prefix3 == "03-" ~ "干豆类及制品",
    prefix3 == "04-" ~ "蔬菜类及制品",
    prefix3 == "05-" ~ "菌藻类",
    prefix3 == "06-" ~ "水果类及制品",
    prefix3 == "07-" ~ "坚果、种子类",
    prefix3 == "08-" ~ "畜肉类及制品",
    prefix3 == "09-" ~ "禽肉类及制品",
    prefix3 == "10-" ~ "乳类及制品",
    prefix3 == "11-" ~ "蛋类及制品",
    prefix3 == "12-" ~ "鱼虾蟹贝类",
    prefix3 == "13-" ~ "婴幼儿食品",
    prefix3 == "14-" ~ "小吃、甜饼类",
    prefix3 == "15-" ~ "速食食品类",
    prefix3 == "16-" ~ "饮料类",
    prefix3 == "17-" ~ "含酒精饮料",
    prefix3 == "18-" ~ "糖、蜜饯类",
    prefix3 == "19-" ~ "油脂类",
    prefix3 == "20-" ~ "调味品类",
    prefix3 == "21-" ~ "药食两用食物及其他类",
    TRUE ~ NA_character_
  )
}

## 1.5 将指定列转 numeric（先 clean_na_char，避免 "NA" -> NaN）
force_numeric_cols <- function(df, num_cols) {
  keep <- intersect(num_cols, names(df))
  if (length(keep) == 0) return(df)
  df %>%
    mutate(across(all_of(keep), ~ suppressWarnings(as.numeric(clean_na_char(.x)))))
}

###############################################################
## 2) 读取 FCT 三个版本 + 清洗 + 生成 food_category（修正版）
###############################################################

## ——统一：从任意编码中提取 2002/2004 的两位大类，并标准化成 "01-" 形式
extract_prefix2_to_01dash <- function(x) {
  x <- clean_na_char(as.character(x))
  x <- str_replace_all(x, "\\s+", "")  # 去空格
  x <- str_replace_all(x, "[^0-9A-Za-z-]", "") # 保留字母数字和-
  
  # 情况1：原本就有 "01-" 这样的形式
  p <- ifelse(str_detect(x, "^\\d{2}-"), str_sub(x, 1, 3), NA_character_)
  
  # 情况2：没有 "-"，但前两位是数字（如 010101 / 01xxxx）
  p2 <- ifelse(is.na(p) & str_detect(x, "^\\d{2}"), paste0(str_sub(x, 1, 2), "-"), NA_character_)
  
  dplyr::coalesce(p, p2)
}

## ——统一：从 1991 编码中提取 "A01"（兼容 A010101、A01-xxx、a01 等）
extract_Axx <- function(x) {
  x <- clean_na_char(as.character(x))
  x <- toupper(x)
  str_extract(x, "A\\d{2}")
}

fct_1991_raw <- read_excel(path_fct, sheet = 1) %>%
  mutate(
    ## 强制把关键编码列转字符并清洗
    foodcode_raw_chr = clean_na_char(as.character(foodcode_raw)),
    foodcode = normalize_foodcode(foodcode),
    ## 提取 A01~A28
    fct_bigcode_1991 = extract_Axx(foodcode_raw_chr),
    food_category = map_cat_1991(fct_bigcode_1991)
  ) %>%
  select(-foodcode_raw_chr)

fct_2002_raw <- read_excel(path_fct, sheet = 2) %>%
  mutate(
    foodcode_raw_chr = clean_na_char(as.character(foodcode_raw)),
    foodcode = normalize_foodcode(foodcode),
    ## 提取并标准化为 "01-" 这种形式（兼容无短横线）
    prefix3 = extract_prefix2_to_01dash(foodcode_raw_chr),
    food_category = map_cat_2002(prefix3)
  ) %>%
  select(-foodcode_raw_chr, -prefix3)

fct_2004_raw <- read_excel(path_fct, sheet = 3) %>%
  mutate(
    foodcode_raw_chr = clean_na_char(as.character(foodcode_raw)),
    foodcode = normalize_foodcode(foodcode),
    prefix3 = extract_prefix2_to_01dash(foodcode_raw_chr),
    food_category = map_cat_2002(prefix3)
  ) %>%
  select(-foodcode_raw_chr, -prefix3)

## ——立刻做一个最小 QC：看 food_category 还有没有全 NA
cat("\n[QC] 1991 food_category NA rate = ",
    mean(is.na(fct_1991_raw$food_category)), "\n")
cat("[QC] 2002 food_category NA rate = ",
    mean(is.na(fct_2002_raw$food_category)), "\n")
cat("[QC] 2004 food_category NA rate = ",
    mean(is.na(fct_2004_raw$food_category)), "\n")

## 抽样看看 foodcode_raw 的真实长相（非常有用）
cat("\n[QC] 示例 foodcode_raw（1991）:\n")
print(head(fct_1991_raw %>% select(foodcode_raw, fct_bigcode_1991, food_category), 10))

cat("\n[QC] 示例 foodcode_raw（2002）:\n")
print(head(fct_2002_raw %>% select(foodcode_raw, food_category), 10))

cat("\n[QC] 示例 foodcode_raw（2004）:\n")
print(head(fct_2004_raw %>% select(foodcode_raw, food_category), 10))

###############################################################
## 3) 只保留“三表共同营养素变量”（列名交集）
##    + 强制保留 food_category（但绝对不能转 numeric）
###############################################################

## 三表共同字段（用于共同营养素列）
common_cols_all3 <- Reduce(
  intersect,
  list(names(fct_1991_raw), names(fct_2002_raw), names(fct_2004_raw))
)

## 关键列（一般三表都有）
must_have <- c("foodcode_raw", "foodcode", "food_name")
must_have <- intersect(must_have, common_cols_all3)

## 共同营养素列：严格三表交集（不含关键列 + 不含 food_category）
nutr_cols_all3 <- setdiff(common_cols_all3, c(must_have, "food_category"))

## 最终保留列：关键列 + 共同营养素列 + food_category
final_keep_cols <- unique(c(must_have, nutr_cols_all3, "food_category"))

## 实际可用列（防止某些列名不存在）
final_keep_cols_1991 <- intersect(final_keep_cols, names(fct_1991_raw))
final_keep_cols_2002 <- intersect(final_keep_cols, names(fct_2002_raw))
final_keep_cols_2004 <- intersect(final_keep_cols, names(fct_2004_raw))

## 只把“共同营养素列”转 numeric（food_category 绝不转）
num_cols <- nutr_cols_all3

fct_1991 <- fct_1991_raw %>%
  select(all_of(final_keep_cols_1991)) %>%
  force_numeric_cols(num_cols) %>%
  mutate(
    food_category = as.character(food_category),
    fct_version_used = "FCT1991"
  )

fct_2002 <- fct_2002_raw %>%
  select(all_of(final_keep_cols_2002)) %>%
  force_numeric_cols(num_cols) %>%
  mutate(
    food_category = as.character(food_category),
    fct_version_used = "FCT2002"
  )

fct_2004 <- fct_2004_raw %>%
  select(all_of(final_keep_cols_2004)) %>%
  force_numeric_cols(num_cols) %>%
  mutate(
    food_category = as.character(food_category),
    fct_version_used = "FCT2004"
  )

###############################################################
## 4) wave>=2004：FCT2002 优先 + FCT2004 兜底（逐字段 coalesce）
##    注意：coalesce 仅在“共同营养素列 + 关键列 + food_category”上进行
###############################################################

## 对 2002/2004 都存在的列进行 coalesce（必须同时存在才安全）
cols_2002 <- names(fct_2002)
cols_2004 <- names(fct_2004)
coalesce_cols <- intersect(cols_2002, cols_2004)
coalesce_cols <- setdiff(coalesce_cols, c("foodcode", "fct_version_used"))

fct_2002_2004 <- fct_2002 %>%
  select(all_of(c("foodcode", coalesce_cols))) %>%
  full_join(
    fct_2004 %>%
      select(all_of(c("foodcode", coalesce_cols))) %>%
      rename_with(~ paste0(.x, "_2004"), -foodcode),
    by = "foodcode"
  ) %>%
  {
    out <- .
    for (nm in coalesce_cols) {
      nm2 <- paste0(nm, "_2004")
      if (nm2 %in% names(out)) out[[nm]] <- dplyr::coalesce(out[[nm]], out[[nm2]])
    }
    out
  } %>%
  mutate(
    fct_version_used = case_when(
      foodcode %in% fct_2002$foodcode ~ "FCT2002",
      foodcode %in% fct_2004$foodcode ~ "FCT2004",
      TRUE ~ NA_character_
    )
  ) %>%
  ## 保留与 fct_2002 一致的列结构（+ fct_version_used）
  select(all_of(c("foodcode", coalesce_cols, "fct_version_used")))

###############################################################
## 5) 处理个人食物条目 data1（只做匹配所需字段）
###############################################################

data_food <- data1 %>%
  mutate(
    IDind = as.character(IDind),
    wave  = suppressWarnings(as.integer(as.character(wave))),
    vd    = suppressWarnings(as.integer(as.character(vd))),
    foodcode = normalize_foodcode(foodcode)
  ) %>%
  filter(!is.na(IDind), !is.na(wave), !is.na(vd), !is.na(foodcode))

###############################################################
## 6) 分段 join（仅匹配；带回 food_category 与共同营养素变量）
###############################################################

## join 要带回的 FCT 列（不重复 foodcode）
fct_cols_return_1991 <- setdiff(names(fct_1991), "foodcode")
fct_cols_return_2002_2004 <- setdiff(names(fct_2002_2004), "foodcode")

data_pre2004 <- data_food %>%
  filter(wave < 2004) %>%
  left_join(
    fct_1991 %>% select(all_of(c("foodcode", fct_cols_return_1991))),
    by = "foodcode"
  )

data_2004p <- data_food %>%
  filter(wave >= 2004) %>%
  left_join(
    fct_2002_2004 %>% select(all_of(c("foodcode", fct_cols_return_2002_2004))),
    by = "foodcode"
  )

###############################################################
## 6.1 显式生成“食物大类指示变量 food_category”
##     ——保证每一条个人食物记录都有明确大类
###############################################################

data_food_matched <- bind_rows(data_pre2004, data_2004p) %>%
  mutate(
    food_category = as.character(food_category)
  )


###############################################################
## 7) QC：按 wave 的匹配率 + 未匹配 foodcode 清单
##    匹配判定：food_name 不为 NA（更稳健）
###############################################################

qc_match_rate_by_wave <- data_food_matched %>%
  mutate(matched = !is.na(food_name)) %>%
  group_by(wave) %>%
  summarise(
    n_records = n(),
    n_matched = sum(matched),
    match_rate = n_matched / n_records,
    .groups = "drop"
  ) %>%
  arrange(wave)

unmatched_foodcodes <- data_food_matched %>%
  mutate(matched = !is.na(food_name)) %>%
  filter(!matched) %>%
  count(wave, foodcode, sort = TRUE)

print(qc_match_rate_by_wave)
print(head(unmatched_foodcodes, 30))
## QC：确认 food_category 不再是全 NA
data_food_matched %>%
  summarise(
    n = n(),
    n_foodcat_na = sum(is.na(food_category)),
    pct_foodcat_na = mean(is.na(food_category))
  ) %>%
  print()

data_food_matched %>%
  count(wave, food_category, sort = TRUE) %>%
  group_by(wave) %>%
  slice_head(n = 5) %>%
  ungroup() %>%
  print(n = 50)

###############################################################
## 8) 导出（按需打开）
###############################################################
write_xlsx(
  list(
    "data_food_matched" = data_food_matched,
    "qc_match_rate_by_wave" = qc_match_rate_by_wave,
    "unmatched_foodcodes" = unmatched_foodcodes
  ),
  path = "个人食物条目_仅匹配FCT_含食物大类_输出与QC_1.xlsx"
)

message("✅ 完成：data_food_matched（含 food_category）+ QC 已生成（如需导出，取消 write_xlsx 注释）")








# 合并家庭和餐次比例
data1 <- read_excel(
  "E:/CHNS/CHNS_2013/初次整理/最终使用数据/食物变量再次处理/家庭食物量表匹配全匹配/匹配家庭食物code.xlsx",
  sheet = 1
)
data2 <- read_excel("E:/CHNS/CHNS_2013/初次整理/食物匹配和处理/匹配子代餐次是否进食/匹配子代餐次是否进食.xlsx", sheet = 2)

data_merged1 <- data1 %>% left_join(data2, by = c("IDind", "wave"))
cat("合并后行数：", nrow(data_merged1), "\n")
cat("成功匹配人数：", length(unique(data_merged1$IDind)), "\n")

library(openxlsx)
write.xlsx(data_merged1, "E:/CHNS/CHNS_2013/初次整理/最终使用数据/食物变量再次处理/家庭食物量表匹配全匹配/匹配家庭餐份.xlsx")



###############################################################
## 家庭食物重新匹配（CHNS foodcode -> FCT1991/2002/2004）
## ✅ 无 vd 版本（与个人条目脚本区分）
##
## 输入：
## - data1: 家庭-食物条目（纵向），至少包含 wave / foodcode
##   （可包含 IDind 或 household id，但不要求 vd）
## - FCT 汇总表：1991(sheet1) / 2002(sheet2) / 2004(sheet3)
##
## 规则：
## - wave < 2004 : 用 FCT1991
## - wave >= 2004: FCT2002 优先 + FCT2004 兜底（逐字段 coalesce）
## - FCT 表中手动填的 "NA" 视为缺失：统一转成真正 NA
## - 只合并“三张 FCT 都有的营养素变量”（列名交集）
## - 额外加入“食物大类 food_category”（在各自 FCT 内生成，不受交集限制）
##
## 输出对象：
## - data_family_food_matched
## - qc_family_match_rate_by_wave
## - unmatched_family_foodcodes
##
## 可选导出 Excel：保留 write_xlsx（建议文件名改为家庭版）
###############################################################

rm(list = ls())

suppressPackageStartupMessages({
  library(dplyr)
  library(stringr)
  library(readxl)
  library(tidyr)
  library(writexl)
})

###############################################################
## 0) 路径 + 读入家庭食物条目
###############################################################
data1 <- read_excel(
  "E:/CHNS/CHNS_2013/初次整理/最终使用数据/食物变量再次处理/家庭食物量表匹配全匹配/匹配家庭餐份.xlsx",
  sheet = 1
)
stopifnot(exists("data1"))

path_fct <- "E:/CHNS/CHNS_2013/食物成分表/整理好的食物成分表/食物成分表汇总.xls"

###############################################################
## 1) 工具函数
###############################################################

## 1.1 伪缺失 -> NA（含手动 "NA"）
clean_na_char <- function(x) {
  if (!is.character(x)) x <- as.character(x)
  x <- trimws(x)
  x[x %in% c("", " ", ".", "NA","Na","na","N/A","n/a","NULL","null","None","none")] <- NA
  x
}

## 1.2 foodcode 统一函数（避免 11101 / 11101.0 / " 11101 "）
normalize_foodcode <- function(x) {
  x <- clean_na_char(x)
  x <- sub("\\.0+$", "", x)
  x <- gsub("[^0-9]", "", x)
  x <- sub("^0+", "", x)
  x[x == ""] <- NA
  x
}

## 1.3 食物大类映射：1991（A01~A28）
map_cat_1991 <- function(bigcode) {
  dplyr::case_when(
    bigcode == "A01" ~ "谷类及制品",
    bigcode == "A02" ~ "干豆类及制品",
    bigcode == "A03" ~ "鲜豆类",
    bigcode == "A04" ~ "根茎类及制品",
    bigcode == "A05" ~ "嫩茎、叶、花类",
    bigcode == "A06" ~ "瓜类",
    bigcode == "A07" ~ "茄果类",
    bigcode == "A08" ~ "咸菜类",
    bigcode == "A09" ~ "菌藻类",
    bigcode == "A10" ~ "鲜果及干果类",
    bigcode == "A11" ~ "坚果类",
    bigcode == "A12" ~ "畜肉类及制品",
    bigcode == "A13" ~ "禽肉类及制品",
    bigcode == "A14" ~ "乳类及制品",
    bigcode == "A15" ~ "婴儿配方食品及辅助食品",
    bigcode == "A16" ~ "蛋类及制品",
    bigcode == "A17" ~ "鱼类",
    bigcode == "A18" ~ "软体动物类",
    bigcode == "A19" ~ "虾蟹类",
    bigcode == "A20" ~ "油脂类",
    bigcode == "A21" ~ "糕点及小吃类",
    bigcode == "A22" ~ "茶及饮料",
    bigcode == "A23" ~ "酒类",
    bigcode == "A24" ~ "糖及制品",
    bigcode == "A25" ~ "淀粉类及制品",
    bigcode == "A26" ~ "调味品类",
    bigcode == "A27" ~ "药用食物类",
    bigcode == "A28" ~ "杂类",
    TRUE ~ NA_character_
  )
}


## 1.4 食物大类映射：2002 / 2004（01- ~ 21-）
map_cat_2002 <- function(prefix3) {
  dplyr::case_when(
    prefix3 == "01-" ~ "谷类及制品",
    prefix3 == "02-" ~ "薯类及淀粉及制品",
    prefix3 == "03-" ~ "干豆类及制品",
    prefix3 == "04-" ~ "蔬菜类及制品",
    prefix3 == "05-" ~ "菌藻类",
    prefix3 == "06-" ~ "水果类及制品",
    prefix3 == "07-" ~ "坚果、种子类",
    prefix3 == "08-" ~ "畜肉类及制品",
    prefix3 == "09-" ~ "禽肉类及制品",
    prefix3 == "10-" ~ "乳类及制品",
    prefix3 == "11-" ~ "蛋类及制品",
    prefix3 == "12-" ~ "鱼虾蟹贝类",
    prefix3 == "13-" ~ "婴幼儿食品",
    prefix3 == "14-" ~ "小吃、甜饼类",
    prefix3 == "15-" ~ "速食食品类",
    prefix3 == "16-" ~ "饮料类",
    prefix3 == "17-" ~ "含酒精饮料",
    prefix3 == "18-" ~ "糖、蜜饯类",
    prefix3 == "19-" ~ "油脂类",
    prefix3 == "20-" ~ "调味品类",
    prefix3 == "21-" ~ "药食两用食物及其他类",
    TRUE ~ NA_character_
  )
}

## 1.5 指定列强制 numeric（避免 "NA" → NaN）
force_numeric_cols <- function(df, num_cols) {
  keep <- intersect(num_cols, names(df))
  if (length(keep) == 0) return(df)
  df %>%
    mutate(across(all_of(keep), ~ suppressWarnings(as.numeric(clean_na_char(.x)))))
}

###############################################################
## 2) 读取 FCT 三版本 + 清洗 + food_category
###############################################################

## —— 提取 2002/2004 大类前缀，统一为 "01-"
extract_prefix2_to_01dash <- function(x) {
  x <- clean_na_char(as.character(x))
  x <- str_replace_all(x, "\\s+", "")
  x <- str_replace_all(x, "[^0-9A-Za-z-]", "")
  p1 <- ifelse(str_detect(x, "^\\d{2}-"), str_sub(x, 1, 3), NA_character_)
  p2 <- ifelse(is.na(p1) & str_detect(x, "^\\d{2}"),
               paste0(str_sub(x, 1, 2), "-"), NA_character_)
  dplyr::coalesce(p1, p2)
}

## —— 提取 1991 的 A01 ~ A28
extract_Axx <- function(x) {
  x <- clean_na_char(as.character(x))
  x <- toupper(x)
  str_extract(x, "A\\d{2}")
}

### ===== FCT 1991 =====
fct_1991_raw <- read_excel(path_fct, sheet = 1) %>%
  mutate(
    foodcode_raw_chr = clean_na_char(as.character(foodcode_raw)),
    foodcode = normalize_foodcode(foodcode),
    fct_bigcode_1991 = extract_Axx(foodcode_raw_chr),
    food_category = map_cat_1991(fct_bigcode_1991)
  ) %>%
  select(-foodcode_raw_chr)

### ===== FCT 2002 =====
fct_2002_raw <- read_excel(path_fct, sheet = 2) %>%
  mutate(
    foodcode_raw_chr = clean_na_char(as.character(foodcode_raw)),
    foodcode = normalize_foodcode(foodcode),
    prefix3 = extract_prefix2_to_01dash(foodcode_raw_chr),
    food_category = map_cat_2002(prefix3)
  ) %>%
  select(-foodcode_raw_chr, -prefix3)

### ===== FCT 2004 =====
fct_2004_raw <- read_excel(path_fct, sheet = 3) %>%
  mutate(
    foodcode_raw_chr = clean_na_char(as.character(foodcode_raw)),
    foodcode = normalize_foodcode(foodcode),
    prefix3 = extract_prefix2_to_01dash(foodcode_raw_chr),
    food_category = map_cat_2002(prefix3)
  ) %>%
  select(-foodcode_raw_chr, -prefix3)

###############################################################
## 3) 三表共同营养素列（严格交集）
###############################################################

common_cols_all3 <- Reduce(
  intersect,
  list(names(fct_1991_raw), names(fct_2002_raw), names(fct_2004_raw))
)

must_have <- intersect(
  c("foodcode_raw", "foodcode", "food_name"),
  common_cols_all3
)

nutr_cols_all3 <- setdiff(common_cols_all3,
                          c(must_have, "food_category"))

final_keep_cols <- unique(c(must_have, nutr_cols_all3, "food_category"))

fct_1991 <- fct_1991_raw %>%
  select(any_of(final_keep_cols)) %>%
  force_numeric_cols(nutr_cols_all3) %>%
  mutate(
    food_category = as.character(food_category),
    fct_version_used = "FCT1991"
  )

fct_2002 <- fct_2002_raw %>%
  select(any_of(final_keep_cols)) %>%
  force_numeric_cols(nutr_cols_all3) %>%
  mutate(
    food_category = as.character(food_category),
    fct_version_used = "FCT2002"
  )

fct_2004 <- fct_2004_raw %>%
  select(any_of(final_keep_cols)) %>%
  force_numeric_cols(nutr_cols_all3) %>%
  mutate(
    food_category = as.character(food_category),
    fct_version_used = "FCT2004"
  )

###############################################################
## 4) wave ≥ 2004：FCT2002 优先 + FCT2004 兜底
###############################################################

coalesce_cols <- intersect(names(fct_2002), names(fct_2004))
coalesce_cols <- setdiff(coalesce_cols, c("foodcode", "fct_version_used"))

fct_2002_2004 <- fct_2002 %>%
  select(foodcode, all_of(coalesce_cols)) %>%
  full_join(
    fct_2004 %>%
      select(foodcode, all_of(coalesce_cols)) %>%
      rename_with(~ paste0(.x, "_2004"), -foodcode),
    by = "foodcode"
  ) %>%
  {
    out <- .
    for (nm in coalesce_cols) {
      nm2 <- paste0(nm, "_2004")
      out[[nm]] <- dplyr::coalesce(out[[nm]], out[[nm2]])
    }
    out
  } %>%
  mutate(
    fct_version_used = case_when(
      foodcode %in% fct_2002$foodcode ~ "FCT2002",
      foodcode %in% fct_2004$foodcode ~ "FCT2004",
      TRUE ~ NA_character_
    )
  ) %>%
  select(foodcode, all_of(coalesce_cols), fct_version_used)

###############################################################
## 5) 家庭食物条目处理（⚠ 无 vd）
###############################################################

## 自动识别家庭ID（若存在）
hhid_candidates <- c("hhid", "HHID", "household_id", "hh_id", "household")
hhid_use <- intersect(hhid_candidates, names(data1))[1]

data_family_food <- data1 %>%
  mutate(
    wave = suppressWarnings(as.integer(as.character(wave))),
    foodcode = normalize_foodcode(foodcode),
    hhid_std = if (!is.na(hhid_use)) as.character(.data[[hhid_use]]) else NA_character_
  ) %>%
  filter(!is.na(wave), !is.na(foodcode))

###############################################################
## 6) 分段 join（家庭版）
###############################################################

data_family_pre2004 <- data_family_food %>%
  filter(wave < 2004) %>%
  left_join(fct_1991, by = "foodcode")

data_family_2004p <- data_family_food %>%
  filter(wave >= 2004) %>%
  left_join(fct_2002_2004, by = "foodcode")

data_family_food_matched <- bind_rows(
  data_family_pre2004,
  data_family_2004p
) %>%
  mutate(food_category = as.character(food_category))

###############################################################
## 7) QC（家庭版）
###############################################################

qc_family_match_rate_by_wave <- data_family_food_matched %>%
  mutate(matched = !is.na(food_name)) %>%
  group_by(wave) %>%
  summarise(
    n_records = n(),
    n_matched = sum(matched),
    match_rate = n_matched / n_records,
    .groups = "drop"
  ) %>%
  arrange(wave)

unmatched_family_foodcodes <- data_family_food_matched %>%
  filter(is.na(food_name)) %>%
  count(wave, foodcode, sort = TRUE)

print(qc_family_match_rate_by_wave)
print(head(unmatched_family_foodcodes, 30))

###############################################################
## 8) 导出（家庭版，避免与个人版混淆）
###############################################################

write_xlsx(
  list(
    "data_family_food_matched" = data_family_food_matched,
    "qc_family_match_rate_by_wave" = qc_family_match_rate_by_wave,
    "unmatched_family_foodcodes" = unmatched_family_foodcodes
  ),
  path = "家庭食物条目_FCT匹配_含食物大类_QC.xlsx"
)

message("✅ 家庭食物匹配完成（无 vd 版本），已与个人脚本彻底区分")




#
###############################################################
## 处理个人食物条目CHNS Dietary Pipeline (Food-entry -> daily -> wave mean3
## -> density vars -> cumulative average) + Excel multi-sheets
## - Unit harmonization for V39:
##   1997-2000: liang (两) -> grams (g) by *50
##   >=2004: already grams
## - No sex-based QC at this stage (per user instruction)
###############################################################

rm(list = ls())

suppressPackageStartupMessages({
  library(dplyr)
  library(tidyr)
  library(stringr)
  library(readxl)
  library(openxlsx)
})

options(dplyr.summarise.inform = FALSE)

###############################################################
## 0) Paths (EDIT THESE)
###############################################################
in_path  <- "E:/博士课题/CHNS/CHNS_2013/初次整理/最终使用数据/食物变量再次处理/个人食物量表全匹配/个人食物条目_仅匹配FCT_含食物大类_输出与QC_1.xlsx"
in_sheet <- 1
out_xlsx <- "E:/博士课题/CHNS/CHNS_2013/初次整理/最终使用数据/食物变量再次处理/个人食物量表全匹配饮食暴露_全流程_daily_wave_cumavg.xlsx"

###############################################################
## 1) Helper functions
###############################################################
clean_na_char <- function(x) {
  if (is.numeric(x)) return(x)
  x_chr <- trimws(as.character(x))
  bad <- c("NA","na","N/A","n/a",""," ",".","..","-","--","NULL","null","NaN","nan")
  x_chr[x_chr %in% bad] <- NA_character_
  x_chr
}

to_num <- function(x) suppressWarnings(as.numeric(clean_na_char(x)))

# Safe divide
safe_div <- function(a, b) ifelse(is.na(b) | b == 0, NA_real_, a / b)

# Cumulative mean for numeric columns within IDind ordered by wave
cummean_vec <- function(x) {
  # x numeric, may have NA
  # cumulative mean ignoring NA (robust)
  out <- rep(NA_real_, length(x))
  s <- 0
  n <- 0
  for (i in seq_along(x)) {
    if (!is.na(x[i])) {
      s <- s + x[i]
      n <- n + 1
    }
    out[i] <- ifelse(n == 0, NA_real_, s / n)
  }
  out
}

###############################################################
## 2) Read data
###############################################################
df0 <- read_excel(in_path, sheet = in_sheet) %>%
  mutate(
    IDind = as.character(IDind),
    wave  = as.character(wave),
    vd    = as.character(vd)
  )

stopifnot(nrow(df0) > 0)

###############################################################
## 3) Define columns: nutrient density fields (per 100g edible)
##    (Use the exact set you listed; keep order stable.)
###############################################################
nutr_cols <- c(
  "energy_kj","energy_kcal",
  "water_g","protein_g","fat_g","fiber_total_g","carb_g","ash_g",
  "vitA_ugRE","carotene_ug",
  "thiamin_mg","riboflavin_mg","niacin_mg",
  "vitC_mg","vitE_total_mg","vitE_alpha_mg",
  "k_mg","na_mg","ca_mg","mg_mg","fe_mg","mn_mg","zn_mg","cu_mg","p_mg","se_ug"
)

missing_nutr <- setdiff(nutr_cols, names(df0))
if (length(missing_nutr) > 0) {
  stop("Missing nutrient columns in input: ", paste(missing_nutr, collapse = ", "))
}

req_cols <- c("IDind","vd","wave","V39","edible_pct","food_category")
missing_req <- setdiff(req_cols, names(df0))
if (length(missing_req) > 0) {
  stop("Missing required columns in input: ", paste(missing_req, collapse = ", "))
}

###############################################################
## 4) Sheet: 00_food_entry_std
##    - Harmonize V39 into grams (food_g_std)
##    - Compute edible grams (edible_g)
##    - Compute per-entry nutrient intakes for ALL nutrient cols
###############################################################
df_entry <- df0 %>%
  mutate(
    wave_num      = suppressWarnings(as.integer(clean_na_char(wave))),
    V39_num       = to_num(V39),
    edible_pct_num= to_num(edible_pct),
    # Unit harmonization: 1997-2000 are liang; >=2004 grams
    # (If you later confirm 2001/2002 exist, adjust the condition.)
    food_g_std = case_when(
      !is.na(wave_num) & wave_num >= 1997 & wave_num <= 2000 ~ V39_num * 50,
      !is.na(wave_num) & wave_num >= 2004 ~ V39_num,
      TRUE ~ V39_num  # fallback: keep as-is (will be flagged later if needed)
    ),
    edible_g = food_g_std * edible_pct_num / 100
  ) %>%
  # clean nutrient densities into numeric
  mutate(across(all_of(nutr_cols), to_num))

# Compute entry-level nutrient intakes:
# assumption: nutrient density columns are per 100g edible portion.
# intake = edible_g/100 * nutrient_density
for (cc in nutr_cols) {
  new_nm <- paste0(cc, "_intake")
  df_entry[[new_nm]] <- (df_entry$edible_g / 100) * df_entry[[cc]]
}

###############################################################
## 5) Sheet: 01_daily_sum (IDind x wave x vd)
##    - Sum nutrient intakes per day
##    - Sum edible grams per day
##    - Sum edible grams by food_category per day (wide)
###############################################################
intake_cols <- paste0(nutr_cols, "_intake")

df_daily_nutr <- df_entry %>%
  group_by(IDind, wave, vd) %>%
  summarise(
    n_food_items_day = dplyr::n(),
    total_food_g_day = sum(edible_g, na.rm = TRUE),
    across(all_of(intake_cols), ~ sum(.x, na.rm = TRUE)),
    .groups = "drop"
  )

# Rename nutrient intake columns to *_day (remove _intake suffix)
names(df_daily_nutr) <- names(df_daily_nutr) %>%
  str_replace("_intake$", "_day")

# Daily food category grams (edible grams), wide
df_daily_cat <- df_entry %>%
  mutate(food_category = clean_na_char(food_category)) %>%
  filter(!is.na(food_category)) %>%
  group_by(IDind, wave, vd, food_category) %>%
  summarise(cat_g_day = sum(edible_g, na.rm = TRUE), .groups = "drop") %>%
  mutate(food_category = make.names(food_category)) %>%
  pivot_wider(
    names_from  = food_category,
    values_from = cat_g_day,
    values_fill = 0
  )

df_daily <- df_daily_nutr %>%
  left_join(df_daily_cat, by = c("IDind","wave","vd"))

###############################################################
## 6) Sheet: 02_wave_3day_mean (IDind x wave)
##    - Mean across vd within wave
##    - Keep n_days_recorded, total items, etc.
###############################################################
# Identify daily numeric columns to average (all *_day + category columns)
daily_num_cols <- df_daily %>%
  select(-IDind, -wave, -vd) %>%
  names()

df_wave_mean3 <- df_daily %>%
  group_by(IDind, wave) %>%
  summarise(
    n_days_recorded = dplyr::n_distinct(vd),
    n_items_wave    = sum(n_food_items_day, na.rm = TRUE),
    across(all_of(daily_num_cols), ~ mean(.x, na.rm = TRUE)),
    .groups = "drop"
  )

# Rename averaged daily cols: *_day -> *_mean3
names(df_wave_mean3) <- names(df_wave_mean3) %>%
  str_replace("_day$", "_mean3")

###############################################################
## 7) Sheet: 03_wave_density
##    - Derive energy shares and density variables from mean3
###############################################################
# Must exist after rename:
stopifnot(all(c("energy_kcal_mean3","protein_g_mean3","fat_g_mean3","carb_g_mean3","fiber_total_g_mean3","total_food_g_mean3") %in%
                names(df_wave_mean3)))

df_wave_density <- df_wave_mean3 %>%
  transmute(
    IDind, wave,
    energy_kcal_mean3,
    protein_g_mean3,
    fat_g_mean3,
    carb_g_mean3,
    fiber_total_g_mean3,
    total_food_g_mean3,
    fat_pct_energy     = safe_div(fat_g_mean3 * 9,  energy_kcal_mean3) * 100,
    protein_pct_energy = safe_div(protein_g_mean3 * 4, energy_kcal_mean3) * 100,
    carb_pct_energy    = safe_div(carb_g_mean3 * 4, energy_kcal_mean3) * 100,
    fiber_g_per_1000kcal = safe_div(fiber_total_g_mean3, energy_kcal_mean3) * 1000,
    energy_density_kcal_per_g = safe_div(energy_kcal_mean3, total_food_g_mean3)
  )

###############################################################
## 8) Sheet: 04_cumavg (IDind x wave)
##    - Cumulative averages for ALL wave-level dietary variables
##      (mean3 variables + density vars)
###############################################################
# Prepare analysis table at wave-level first (merge mean3 + density)
df_wave_all <- df_wave_mean3 %>%
  left_join(df_wave_density %>% select(-energy_kcal_mean3, -protein_g_mean3, -fat_g_mean3, -carb_g_mean3, -fiber_total_g_mean3, -total_food_g_mean3),
            by = c("IDind","wave")) %>%
  mutate(wave_num = suppressWarnings(as.integer(clean_na_char(wave)))) %>%
  arrange(IDind, wave_num, wave)

# Cumulative mean for ALL numeric dietary columns except identifiers
id_cols <- c("IDind","wave","wave_num")
diet_num_cols <- df_wave_all %>%
  select(-all_of(id_cols)) %>%
  select(where(is.numeric)) %>%
  names()

df_cumavg <- df_wave_all %>%
  group_by(IDind) %>%
  arrange(wave_num, .by_group = TRUE) %>%
  mutate(
    across(all_of(diet_num_cols), ~ cummean_vec(.x), .names = "cum_{.col}")
  ) %>%
  ungroup() %>%
  select(IDind, wave, all_of(paste0("cum_", diet_num_cols)))

###############################################################
## 9) Sheet: 05_analysis_ready_diet
##    - A convenient final table (mean3 + density + cumavg)
###############################################################
df_analysis_ready <- df_wave_all %>%
  select(-wave_num) %>%
  left_join(df_cumavg, by = c("IDind","wave"))

###############################################################
## 10) Write all sheets to one Excel
###############################################################
wb <- createWorkbook()

addWorksheet(wb, "00_food_entry_std")
writeData(wb, "00_food_entry_std", df_entry)

addWorksheet(wb, "01_daily_sum")
writeData(wb, "01_daily_sum", df_daily)

addWorksheet(wb, "02_wave_3day_mean")
writeData(wb, "02_wave_3day_mean", df_wave_mean3)

addWorksheet(wb, "03_wave_density")
writeData(wb, "03_wave_density", df_wave_density)

addWorksheet(wb, "04_cumavg")
writeData(wb, "04_cumavg", df_cumavg)

addWorksheet(wb, "05_analysis_ready_diet")
writeData(wb, "05_analysis_ready_diet", df_analysis_ready)

saveWorkbook(wb, out_xlsx, overwrite = TRUE)

message("✅ Finished. Output saved to: ", out_xlsx)



library(readxl)
library(dplyr)
library(openxlsx)

# 读入数据
data1 <- read_excel(
  "E:/CHNS/CHNS_2013/初次整理/最终使用数据/食物变量再次处理/个人食物量表全匹配/调整城市化指数变量.xlsx",
  sheet = 1
)

data2 <- read_excel(
  "E:/CHNS/CHNS_2013/初次整理/最终使用数据/食物变量再次处理/个人食物量表全匹配/个人食物量表全匹配饮食暴露_全流程_daily_wave_cumavg.xlsx",
  sheet = 6
)

# 统一 join key 类型
data1 <- data1 %>%
  mutate(
    IDind = as.character(IDind),
    wave  = as.character(wave)
  )

data2 <- data2 %>%
  mutate(
    IDind = as.character(IDind),
    wave  = as.character(wave)
  )

# 正确合并（推荐）
data_merged1 <- data1 %>%
  left_join(data2, by = c("IDind", "wave"))

# 基本检查
stopifnot(nrow(data_merged1) == nrow(data1))
cat("合并后行数：", nrow(data_merged1), "\n")

# 导出
write.xlsx(
  data_merged1,
  "E:/CHNS/CHNS_2013/初次整理/最终使用数据/食物变量再次处理/个人食物量表全匹配/finaldata.xlsx",
  overwrite = TRUE
)




###############################################################
## FINAL_FAST: 孕前 BMI -> 子代饮食摄入（多结局批量回归）
## - LMM: y ~ BMI + covars + (1|IDind)
## - 食物组结局：自动加入能量（energy_kcal_mean3 / cum_energy_kcal_mean3）
## - 输出：Excel（各暴露一个sheet）+ Forest plots（每暴露×类别一张）
##
## 协变量（按你定义）：
## wave, SEX_2, han, age, A5C, AGE_m, V29_m, A12_m,
## hhsize, region_ns_num (1/2), index_tertile_num (1/2/3), log_hhincpc
###############################################################

rm(list = ls())

suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(stringr)
  library(tidyr)
  library(purrr)
  library(ggplot2)
  library(lme4)
  library(lmerTest)
  library(broom.mixed)
  library(openxlsx)
})

options(dplyr.summarise.inform = FALSE)

###############################################################
## 0) 路径与读入
###############################################################
path_data <- "E:/CHNS/CHNS_2013/初次整理/最终使用数据/食物变量再次处理/个人食物量表全匹配/finaldata_pca_UPF 食物占比.xlsx"

dir_base <- "E:/CHNS/CHNS_2013/初次整理/最终使用数据/食物变量再次处理/个人食物量表全匹配/05_BMI_to_diet_outputs_FINAL_FAST"
dir_fig  <- file.path(dir_base, "fig_forest")
dir_tab  <- file.path(dir_base, "tables")
dir.create(dir_fig, recursive = TRUE, showWarnings = FALSE)
dir.create(dir_tab, recursive = TRUE, showWarnings = FALSE)

data0 <- read_excel(path_data, sheet = 1) %>%
  mutate(
    IDind = as.character(IDind),
    wave  = as.character(wave)
  )
stopifnot(nrow(data0) > 0)

###############################################################
## 1) 工具函数：伪缺失 + numeric
###############################################################
clean_na_char <- function(x) {
  x <- trimws(as.character(x))
  x[x %in% c("", " ", ".", "NA", "na", "Na", "N/A", "-", "--")] <- NA
  x
}
clean_num <- function(x) {
  if (is.numeric(x)) return(x)
  x <- clean_na_char(x)
  suppressWarnings(as.numeric(x))
}

# 两列求和：两者都 NA -> NA，否则缺失当 0
safe_sum2 <- function(a, b) {
  a0 <- clean_num(a)
  b0 <- clean_num(b)
  both_na <- is.na(a0) & is.na(b0)
  out <- coalesce(a0, 0) + coalesce(b0, 0)
  out[both_na] <- NA_real_
  out
}

###############################################################
## 2) 预处理：一次性清洗字符、合并变量、类型设定
###############################################################
data1 <- data0 %>%
  mutate(across(where(is.character), clean_na_char)) %>%
  mutate(
    ## 合并“高糖”(g/day)
    high_sugar_g_mean3        = if (all(c("糖.蜜饯类","糖及制品") %in% names(.))) safe_sum2(`糖.蜜饯类`, `糖及制品`) else NA_real_,
    dessert_snack_g_mean3     = if (all(c("糕点及小吃类","小吃.甜饼类") %in% names(.))) safe_sum2(`糕点及小吃类`, `小吃.甜饼类`) else NA_real_,
    
    ## 合并 cum 版本(g/day)
    cum_high_sugar_g_mean3    = if (all(c("cum_糖.蜜饯类","cum_糖及制品") %in% names(.))) safe_sum2(`cum_糖.蜜饯类`, `cum_糖及制品`) else NA_real_,
    cum_dessert_snack_g_mean3 = if (all(c("cum_糕点及小吃类","cum_小吃.甜饼类") %in% names(.))) safe_sum2(`cum_糕点及小吃类`, `cum_小吃.甜饼类`) else NA_real_
  )

## 类型设定：按你定义
fac_vars <- intersect(
  c("wave","SEX_2","han","A5C","V29_m","A12_m","region_ns_num","index_tertile_num"),
  names(data1)
)
data1[fac_vars] <- lapply(data1[fac_vars], function(z) droplevels(as.factor(clean_na_char(z))))

num_vars <- intersect(c("age","AGE_m","hhsize","log_hhincpc"), names(data1))
for (v in num_vars) data1[[v]] <- clean_num(data1[[v]])

# 暴露：孕前 BMI（连续 + 二分类版本可选）
exposures <- c("bmi_m_pre", "bmi_m_pre_bin24_num", "bmi_m_pre_bin25_num")
exposures <- exposures[exposures %in% names(data1)]
if (length(exposures) == 0) stop("未在数据中找到孕前 BMI 暴露列，请检查列名。")
for (x in exposures) data1[[x]] <- clean_num(data1[[x]])

# 能量列（供能量校正）
if ("energy_kcal_mean3" %in% names(data1)) data1$energy_kcal_mean3 <- clean_num(data1$energy_kcal_mean3)
if ("cum_energy_kcal_mean3" %in% names(data1)) data1$cum_energy_kcal_mean3 <- clean_num(data1$cum_energy_kcal_mean3)

###############################################################
## 3) 协变量（严格按你的清单；存在才纳入）
###############################################################
covars_fixed <- c(
  "wave","SEX_2","han","age",
  "A5C","AGE_m","V29_m","A12_m",
  "hhsize","region_ns_num","index_tertile_num","log_hhincpc"
)
covars <- intersect(covars_fixed, names(data1))

cov_missing <- setdiff(covars_fixed, names(data1))
if (length(cov_missing) > 0) {
  message("⚠️ 以下协变量在数据中不存在，已自动跳过：", paste(cov_missing, collapse = ", "))
}

###############################################################
## 4) 结局变量集合（尽量都跑）
###############################################################
# wave-level 营养素/总量
outcomes_wave_nutr <- intersect(c(
  "total_food_g_mean3","energy_kj_mean3","energy_kcal_mean3","water_g_mean3",
  "protein_g_mean3","fat_g_mean3","carb_g_mean3","fiber_total_g_mean3","ash_g_mean3",
  "vitA_ugRE_mean3","carotene_ug_mean3","thiamin_mg_mean3","riboflavin_mg_mean3",
  "niacin_mg_mean3","vitC_mg_mean3","vitE_total_mg_mean3","vitE_alpha_mg_mean3",
  "k_mg_mean3","na_mg_mean3","ca_mg_mean3","mg_mg_mean3","fe_mg_mean3","mn_mg_mean3",
  "zn_mg_mean3","cu_mg_mean3","p_mg_mean3","se_ug_mean3"
), names(data1))

# wave-level 结构指标
outcomes_wave_struct <- intersect(c(
  "fat_pct_energy","protein_pct_energy","carb_pct_energy",
  "fiber_g_per_1000kcal","energy_density_kcal_per_g"
), names(data1))

# wave-level 食物组（含合并后的两列）
food_group_wave <- intersect(c(
  "乳类及制品","水果类及制品","薯类及淀粉及制品","蛋类及制品","谷类及制品","速食食品类",
  "鱼虾蟹贝类","禽肉类及制品","蔬菜类及制品","畜肉类及制品","干豆类及制品","菌藻类",
  "调味品类","坚果.种子类","饮料类","油脂类",
  "嫩茎.叶.花类","根茎类及制品","鲜豆类","瓜类","茄果类","鱼类","鲜果及干果类",
  "坚果类","虾蟹类","软体动物类","咸菜类","茶及饮料","淀粉类及制品",
  "婴儿配方食品及辅助食品","婴幼儿食品","酒类","含酒精饮料","杂类",
  "糖.蜜饯类","糖及制品","糕点及小吃类","小吃.甜饼类",
  "high_sugar_g_mean3","dessert_snack_g_mean3"
), names(data1))

# cum-level 营养素/总量
outcomes_cum_nutr <- intersect(c(
  "cum_total_food_g_mean3","cum_energy_kj_mean3","cum_energy_kcal_mean3","cum_water_g_mean3",
  "cum_protein_g_mean3","cum_fat_g_mean3","cum_carb_g_mean3","cum_fiber_total_g_mean3","cum_ash_g_mean3",
  "cum_vitA_ugRE_mean3","cum_carotene_ug_mean3","cum_thiamin_mg_mean3","cum_riboflavin_mg_mean3",
  "cum_niacin_mg_mean3","cum_vitC_mg_mean3","cum_vitE_total_mg_mean3","cum_vitE_alpha_mg_mean3",
  "cum_k_mg_mean3","cum_na_mg_mean3","cum_ca_mg_mean3","cum_mg_mg_mean3","cum_fe_mg_mean3","cum_mn_mg_mean3",
  "cum_zn_mg_mean3","cum_cu_mg_mean3","cum_p_mg_mean3","cum_se_ug_mean3"
), names(data1))

# cum-level 结构指标
outcomes_cum_struct <- intersect(c(
  "cum_fat_pct_energy","cum_protein_pct_energy","cum_cum_carb_pct_energy",
  "cum_fiber_g_per_1000kcal","cum_energy_density_kcal_per_g"
), names(data1))
# 修正可能的拼写错误：cum_cum_carb_pct_energy
outcomes_cum_struct <- setdiff(outcomes_cum_struct, "cum_cum_carb_pct_energy")
outcomes_cum_struct <- unique(c(outcomes_cum_struct, intersect("cum_carb_pct_energy", names(data1))))

# cum-level 食物组（含合并后的两列）
food_group_cum <- paste0("cum_", c(
  "乳类及制品","水果类及制品","薯类及淀粉及制品","蛋类及制品","谷类及制品","速食食品类",
  "鱼虾蟹贝类","禽肉类及制品","蔬菜类及制品","畜肉类及制品","干豆类及制品","菌藻类",
  "调味品类","坚果.种子类","饮料类","油脂类",
  "嫩茎.叶.花类","根茎类及制品","鲜豆类","瓜类","茄果类","鱼类","鲜果及干果类",
  "坚果类","虾蟹类","软体动物类","咸菜类","茶及饮料","淀粉类及制品",
  "婴儿配方食品及辅助食品","婴幼儿食品","酒类","含酒精饮料","杂类",
  "糖.蜜饯类","糖及制品","糕点及小吃类","小吃.甜饼类"
))
food_group_cum <- intersect(food_group_cum, names(data1))
food_group_cum <- unique(c(food_group_cum, intersect(c("cum_high_sugar_g_mean3","cum_dessert_snack_g_mean3"), names(data1))))

outcome_sets <- list(
  wave_nutrients  = outcomes_wave_nutr,
  wave_structure  = outcomes_wave_struct,
  wave_foodgroups = food_group_wave,
  cum_nutrients   = outcomes_cum_nutr,
  cum_structure   = outcomes_cum_struct,
  cum_foodgroups  = food_group_cum
)

###############################################################
## 5) 建模：FINAL_FAST 函数
###############################################################
is_food_group <- function(y) y %in% c(food_group_wave, food_group_cum)

energy_var_for_outcome <- function(y, dat_names) {
  if (str_starts(y, "cum_")) {
    if ("cum_energy_kcal_mean3" %in% dat_names) return("cum_energy_kcal_mean3")
    return(NA_character_)
  } else {
    if ("energy_kcal_mean3" %in% dat_names) return("energy_kcal_mean3")
    return(NA_character_)
  }
}

fit_lmm_one_fast <- function(dat, y, x, covars) {
  
  # 仅转换 y；x 已预处理为 numeric
  y_vec <- clean_num(dat[[y]])
  x_vec <- dat[[x]]
  
  # 组装协变量
  cov_use <- covars
  evar <- energy_var_for_outcome(y, names(dat))
  if (is_food_group(y) && !is.na(evar) && evar %in% names(dat) && y != evar) {
    cov_use <- unique(c(cov_use, evar))
  }
  cov_use <- cov_use[cov_use %in% names(dat)]
  
  keep_cols <- unique(c("IDind", x, cov_use))
  keep_cols <- intersect(keep_cols, names(dat))
  d <- dat[, keep_cols, drop = FALSE]
  d$y_num <- y_vec
  
  # 完整案例
  ok <- !is.na(d$y_num) & !is.na(d[[x]]) & !is.na(d$IDind)
  d <- d[ok, , drop = FALSE]
  
  if (nrow(d) < 200) return(NULL)
  if (length(unique(d[[x]])) < 2) return(NULL)
  
  rhs <- paste(c(x, cov_use), collapse = " + ")
  fml <- as.formula(paste0("y_num ~ ", rhs, " + (1|IDind)"))
  
  m <- tryCatch(
    suppressWarnings(
      lmer(
        fml, data = d, REML = FALSE,
        control = lmerControl(
          check.rankX = "ignore",
          optimizer = "bobyqa",
          optCtrl = list(maxfun = 1e5)
        )
      )
    ),
    error = function(e) NULL
  )
  if (is.null(m)) return(NULL)
  
  broom.mixed::tidy(m, effects = "fixed", conf.int = TRUE) %>%
    filter(term == x) %>%
    transmute(
      outcome = y,
      exposure = x,
      estimate = estimate,
      conf.low = conf.low,
      conf.high = conf.high,
      p.value = p.value,
      n_obs = nrow(d),
      n_id  = dplyr::n_distinct(d$IDind),
      singular = lme4::isSingular(m, tol = 1e-4)
    )
}

run_batch_fast <- function(dat, exposures, covars, outcome_sets) {
  
  # 只保留必要列（极大提速）
  all_outcomes <- unique(unlist(outcome_sets))
  keep0 <- unique(c("IDind", exposures, covars, "energy_kcal_mean3", "cum_energy_kcal_mean3", all_outcomes))
  keep0 <- intersect(keep0, names(dat))
  dat_small <- dat[, keep0, drop = FALSE]
  
  res_all <- list()
  for (x in exposures) {
    message("=== Running exposure: ", x, " ===")
    tmp <- purrr::imap_dfr(outcome_sets, function(ys, set_name) {
      if (length(ys) == 0) return(NULL)
      rr <- purrr::map_dfr(ys, ~fit_lmm_one_fast(dat_small, y = .x, x = x, covars = covars))
      if (is.null(rr) || nrow(rr) == 0) return(NULL)
      rr %>% mutate(outcome_set = set_name)
    })
    res_all[[x]] <- tmp
  }
  res_all
}

res_list <- run_batch_fast(data1, exposures, covars, outcome_sets)

###############################################################
## 6) 导出 Excel（每个 exposure 一个 sheet）
###############################################################
path_xlsx <- file.path(dir_tab, "results_BMI_to_child_diet_LMM_FINAL_FAST.xlsx")
wb <- createWorkbook()

for (x in names(res_list)) {
  df <- res_list[[x]]
  if (is.null(df) || nrow(df) == 0) next
  
  df2 <- df %>%
    mutate(
      beta_CI = sprintf("%.4f (%.4f, %.4f)", estimate, conf.low, conf.high),
      p_fmt   = ifelse(is.na(p.value), NA_character_,
                       ifelse(p.value < 0.001, "<0.001", sprintf("%.3f", p.value)))
    ) %>%
    arrange(outcome_set, p.value)
  
  sh <- str_sub(x, 1, 28)
  addWorksheet(wb, sheetName = sh)
  writeData(wb, sheet = sh, x = df2)
}

saveWorkbook(wb, path_xlsx, overwrite = TRUE)
message("✅ Excel saved: ", path_xlsx)

###############################################################
## 7) Forest plots：每 exposure × outcome_set 一张（最多显示前40个结局）
###############################################################
make_forest <- function(df, title) {
  df <- df %>% filter(!is.na(estimate), !is.na(conf.low), !is.na(conf.high))
  if (nrow(df) == 0) return(NULL)
  
  df <- df %>%
    arrange(p.value) %>%
    mutate(outcome = factor(outcome, levels = rev(unique(outcome))))
  
  ggplot(df, aes(x = estimate, y = outcome)) +
    geom_vline(xintercept = 0, linetype = 2) +
    geom_point(size = 2) +
    geom_errorbarh(aes(xmin = conf.low, xmax = conf.high), height = 0.2) +
    labs(
      title = title,
      x = "Beta (change in diet outcome per 1-unit increase in exposure)",
      y = NULL
    ) +
    theme_bw(base_size = 11)
}

for (x in names(res_list)) {
  df <- res_list[[x]]
  if (is.null(df) || nrow(df) == 0) next
  
  for (set_name in unique(df$outcome_set)) {
    dfi <- df %>% filter(outcome_set == set_name) %>% arrange(p.value)
    dfi_plot <- dfi %>% slice(1:min(40, n()))
    
    p <- make_forest(
      dfi_plot,
      title = paste0("Pre-pregnancy BMI -> child diet outcomes (", x, " | ", set_name, ")")
    )
    if (is.null(p)) next
    
    fn <- file.path(dir_fig, paste0("forest_", x, "_", set_name, ".png"))
    ggsave(fn, p, width = 10, height = 8, dpi = 300)
  }
}
message("✅ Forest plots saved in: ", dir_fig)

###############################################################
## 8) QC summary（可选）
###############################################################
qc <- tibble::tibble(
  n_obs = nrow(data1),
  n_id  = dplyr::n_distinct(data1$IDind),
  n_wave = dplyr::n_distinct(data1$wave),
  exposure_vars = paste(exposures, collapse = ", "),
  covars_used   = paste(covars, collapse = ", ")
)
write.xlsx(qc, file.path(dir_tab, "QC_summary_FINAL_FAST.xlsx"), overwrite = TRUE)
message("✅ QC saved.")







# 新思路再次分析
# ============================================================
# CHNS finaldata: Continuous outcomes (LMM) + P75 binary (GEE)
# FULLY-ADJUSTED models with consistent covariate set
# - LMM: lmerTest::lmer (REML=FALSE)
# - GEE: geepack::geeglm (binomial, exchangeable)
# - P75 cutoff: within (wave × sex_str)
# - Strata: Overall + Sex + Age
# - Output: 2 Excel files
# ============================================================

rm(list = ls())

suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(stringr)
  library(lme4)
  library(lmerTest)
  library(broom.mixed)
  library(geepack)
  library(broom)
  library(openxlsx)
  library(tidyr)
})

options(dplyr.summarise.inform = FALSE)

# ------------------------------------------------------------
# 0) 数据导入 + 输出目录
# ------------------------------------------------------------
data_path <- "E:/博士课题/CHNS/CHNS_2013/初次整理/最终使用数据/食物变量再次处理/个人食物量表全匹配/finaldata.xlsx"
if (!file.exists(data_path)) stop("文件路径错误：", data_path)

out_dir <- file.path(dirname(data_path), "FULLADJ_LMM_GEE_outputs")
dir.create(out_dir, showWarnings = FALSE, recursive = TRUE)

raw0 <- readxl::read_excel(data_path, sheet = 1)

# ------------------------------------------------------------
# 1) 变量定义
# ------------------------------------------------------------
outcomes <- c(
  "fat_pct_energy", "protein_pct_energy", "carb_pct_energy",
  "fiber_g_per_1kcal", "energy_density_kcal_per_g",
  "cum_energy_kj_mean3", "cum_energy_kcal_mean3", "cum_fat_pct_energy",
  "cum_protein_pct_energy", "cum_carb_pct_energy", "cum_fiber_g_per_1kcal",
  "cum_energy_density_kcal_per_g", "energy_kj_mean3", "energy_kcal_mean3",
  "protein_g_mean3", "fat_g_mean3", "fiber_total_g_mean3", "carb_g_mean3"
)

# FULL adjustment set (一致的协变量集合)
covars_full <- c(
  "wave", "SEX_2", "han", "age", "A5C", "AGE_m", "V29_m", "A12_m",
  "hhsize", "region_ns_num", "index_tertile_num", "log_hhincpc"
)

# ------------------------------------------------------------
# 2) 工具函数：安全数值化 + 标准化
# ------------------------------------------------------------
to_num <- function(x) {
  if (is.factor(x)) x <- as.character(x)
  if (is.character(x)) {
    x <- str_replace_all(x, ",", "")
    x <- str_trim(x)
  }
  suppressWarnings(as.numeric(x))
}

zscore <- function(x) {
  if (all(is.na(x))) return(x)
  s <- sd(x, na.rm = TRUE)
  if (is.na(s) || s == 0) return(x)
  as.numeric(scale(x))
}

# 子集内剔除“常数/无变异”协变量（避免模型报错）
drop_invariant_covars <- function(dat, covars) {
  keep <- c()
  for (v in covars) {
    if (!v %in% names(dat)) next
    x <- dat[[v]]
    if (is.factor(x) || is.character(x)) {
      x2 <- droplevels(as.factor(x))
      if (nlevels(x2) >= 2) keep <- c(keep, v)
    } else {
      if (sum(!is.na(x)) >= 2 && sd(x, na.rm = TRUE) > 0) keep <- c(keep, v)
    }
  }
  unique(keep)
}

fmt_p <- function(p) {
  if (is.na(p)) return(NA_character_)
  if (p < 0.001) return("<0.001")
  sprintf("%.3f", p)
}

# ------------------------------------------------------------
# 3) 预处理：类型修复 + 分层变量
# ------------------------------------------------------------
data1 <- raw0 %>%
  mutate(
    IDind = as.factor(IDind),
    
    bmi_m_pre   = to_num(bmi_m_pre),
    age         = to_num(age),
    AGE_m       = if ("AGE_m" %in% names(.)) to_num(AGE_m) else NA_real_,
    V29_m       = if ("V29_m" %in% names(.)) to_num(V29_m) else NA_real_,
    A12_m       = if ("A12_m" %in% names(.)) to_num(A12_m) else NA_real_,
    hhsize      = if ("hhsize" %in% names(.)) to_num(hhsize) else NA_real_,
    log_hhincpc = if ("log_hhincpc" %in% names(.)) to_num(log_hhincpc) else NA_real_,
    
    age_group = case_when(
      !is.na(age) & age <= 6              ~ "0-6",
      !is.na(age) & age >= 7 & age <= 10  ~ "7-10",
      !is.na(age) & age > 10              ~ ">10",
      TRUE ~ NA_character_
    ),
    sex_str = case_when(
      SEX_2 == 1 ~ "Male",
      SEX_2 == 2 ~ "Female",
      TRUE ~ NA_character_
    )
  )

# 因子变量统一为 factor
factor_vars <- intersect(
  c("wave","SEX_2","han","A5C","region_ns_num","index_tertile_num","sex_str","age_group"),
  names(data1)
)
data1 <- data1 %>% mutate(across(all_of(factor_vars), as.factor))

# 结局变量转数值
outcomes_exist <- intersect(outcomes, names(data1))
data1 <- data1 %>% mutate(across(all_of(outcomes_exist), to_num))

# ------------------------------------------------------------
# 4) LMM：全调整协变量（与 GEE 一致）
# ------------------------------------------------------------
fit_lmm_extract <- function(dat, y, stratum_label, strat_type) {
  
  if (!y %in% names(dat)) return(NULL)
  
  cov_use <- covars_full[covars_full %in% names(dat)]
  # Sex 分层：移除 SEX_2
  if (strat_type == "Sex") cov_use <- setdiff(cov_use, "SEX_2")
  
  need_vars <- unique(c(y, "bmi_m_pre", "IDind", cov_use))
  dat_model <- dat %>% select(all_of(need_vars)) %>% drop_na()
  
  if (nrow(dat_model) < 30) return(NULL)
  
  # 子集内剔除无变异协变量（避免奇异）
  cov_use2 <- drop_invariant_covars(dat_model, cov_use)
  
  # 标准化数值协变量（提升收敛）
  num_covs <- intersect(
    c("bmi_m_pre","age","AGE_m","V29_m","A12_m","hhsize","log_hhincpc"),
    names(dat_model)
  )
  dat_model <- dat_model %>% mutate(across(all_of(num_covs), zscore))
  
  # 防御：字符转因子
  for (v in cov_use2) {
    if (v %in% names(dat_model) && is.character(dat_model[[v]])) {
      dat_model[[v]] <- as.factor(dat_model[[v]])
    }
  }
  
  rhs <- c("bmi_m_pre", cov_use2)
  fml <- as.formula(paste0(y, " ~ ", paste(rhs, collapse = " + "), " + (1 | IDind)"))
  
  m <- tryCatch(
    lmerTest::lmer(
      fml, data = dat_model, REML = FALSE,
      control = lmerControl(optimizer = "bobyqa", optCtrl = list(maxfun = 2e5))
    ),
    error = function(e) NULL
  )
  if (is.null(m)) return(NULL)
  
  td <- broom.mixed::tidy(m, effects = "fixed", conf.int = TRUE, conf.method = "Wald")
  has_p <- "p.value" %in% names(td)
  if (!has_p) td$p.value <- NA_real_
  
  out <- td %>%
    filter(term == "bmi_m_pre") %>%
    mutate(
      outcome = y,
      stratum_type = strat_type,
      stratum = stratum_label,
      n_obs = nobs(m),
      n_id  = dplyr::n_distinct(dat_model$IDind),
      is_singular = isSingular(m),
      converged = is.null(m@optinfo$conv$lme4$messages),
      covars_used = if (length(cov_use2) == 0) "" else paste(cov_use2, collapse = " + "),
      p_note = if (!has_p) "p.value not returned by tidy(); left as NA" else ""
    ) %>%
    transmute(
      outcome, stratum_type, stratum,
      n_obs, n_id,
      estimate, conf.low, conf.high, std.error, statistic, p.value,
      is_singular, converged,
      covars_used, p_note
    ) %>%
    mutate(
      beta_ci = sprintf("%.3f (%.3f, %.3f)", estimate, conf.low, conf.high),
      p_show  = vapply(p.value, fmt_p, character(1))
    )
  
  out
}

# ------------------------------------------------------------
# 5) GEE：P75 二分类（wave × sex 内阈值）+ 全调整协变量
# ------------------------------------------------------------

# 在 wave × sex_str 内计算 P75；若该组可用 y 太少，则 cutoff 为 NA
make_p75_bin <- function(dat, y, p = 0.75, min_n = 10) {
  dat %>%
    group_by(wave, sex_str) %>%
    mutate(
      n_nonmiss = sum(!is.na(.data[[y]])),
      cutoff_p = ifelse(n_nonmiss >= min_n,
                        as.numeric(quantile(.data[[y]], probs = p, na.rm = TRUE)),
                        NA_real_),
      y_bin = ifelse(!is.na(cutoff_p) & !is.na(.data[[y]]) & .data[[y]] >= cutoff_p, 1L,
                     ifelse(!is.na(cutoff_p) & !is.na(.data[[y]]), 0L, NA_integer_))
    ) %>%
    ungroup() %>%
    select(-n_nonmiss)
}

fit_gee_extract <- function(dat, y, stratum_type, stratum_label) {
  
  if (!y %in% names(dat)) return(NULL)
  
  # 全调整协变量（与 LMM 同一集合），Sex 分层移除 SEX_2
  cov_use <- covars_full[covars_full %in% names(dat)]
  if (stratum_type == "Sex") cov_use <- setdiff(cov_use, "SEX_2")
  
  # 为了计算 wave×sex 的 P75，需要 wave + sex_str（即便 Sex 分层也需要 wave；sex_str 在子集中是常量也没关系）
  need_for_cut <- intersect(c("wave", "sex_str"), names(dat))
  
  # 模型需要的变量（y_bin 在 make_p75_bin 后生成）
  need_vars <- unique(c("IDind", "bmi_m_pre", cov_use, need_for_cut))
  
  # 先在完整 dat 上生成 y_bin（只要 wave & sex_str 存在）
  dat_y <- dat
  if (!all(c("wave","sex_str") %in% names(dat_y))) return(NULL)
  dat_y <- make_p75_bin(dat_y, y, p = 0.75, min_n = 10)
  
  # 组装模型数据：y_bin + 所需协变量 + 暴露
  dat_model <- dat_y %>%
    select(all_of(unique(c("IDind", "y_bin", "bmi_m_pre", cov_use)))) %>%
    drop_na()
  
  if (nrow(dat_model) < 100) return(NULL)
  if (length(unique(dat_model$y_bin)) < 2) return(NULL)
  
  # 子集内剔除无变异协变量（避免 geeglm 因单水平因子报错）
  cov_use2 <- drop_invariant_covars(dat_model, cov_use)
  
  # 去掉未使用因子水平（geeglm 常见要求）
  dat_model <- dat_model %>% mutate(across(where(is.factor), droplevels))
  
  # 构建公式
  rhs <- c("bmi_m_pre", cov_use2)
  fml <- as.formula(paste0("y_bin ~ ", paste(rhs, collapse = " + ")))
  
  m <- tryCatch(
    geepack::geeglm(
      fml,
      id = IDind,
      family = binomial(),
      corstr = "exchangeable",
      data = dat_model
    ),
    error = function(e) NULL
  )
  if (is.null(m)) return(NULL)
  
  td <- broom::tidy(m, exponentiate = TRUE, conf.int = TRUE) %>%
    filter(term == "bmi_m_pre")
  
  if (nrow(td) == 0) return(NULL)
  
  out <- td %>%
    mutate(
      outcome = y,
      stratum_type = stratum_type,
      stratum = stratum_label,
      n_obs = nrow(dat_model),
      n_id  = dplyr::n_distinct(dat_model$IDind),
      n_event = sum(dat_model$y_bin == 1),
      event_rate = mean(dat_model$y_bin == 1),
      covars_used = if (length(cov_use2) == 0) "" else paste(cov_use2, collapse = " + "),
      separation_flag = ifelse(
        is.na(estimate) |
          !is.finite(estimate) |
          !is.finite(conf.low) |
          !is.finite(conf.high) |
          estimate == 0 | conf.low == 0 | conf.high == 0,
        TRUE, FALSE
      )
    ) %>%
    transmute(
      outcome, stratum_type, stratum,
      n_obs, n_id, n_event, event_rate,
      estimate, conf.low, conf.high, p.value,
      covars_used, separation_flag
    ) %>%
    mutate(
      OR_CI = sprintf("%.2f (%.2f, %.2f)", estimate, conf.low, conf.high),
      p_show = vapply(p.value, fmt_p, character(1)),
      event_rate_pct = sprintf("%.1f%%", event_rate * 100)
    )
  
  out
}

# ------------------------------------------------------------
# 6) 批量运行：Overall + Sex + Age（两套模型同步）
# ------------------------------------------------------------
res_lmm <- list()
res_gee <- list()

sex_levels <- c("Male", "Female")
age_levels <- c("0-6", "7-10", ">10")

for (y in outcomes) {
  
  # ---------------- LMM ----------------
  res_lmm[[paste0(y, "_Overall_All")]] <- fit_lmm_extract(data1, y, "All", "Overall")
  
  for (s in sex_levels) {
    sub <- data1 %>% filter(sex_str == s)
    res_lmm[[paste0(y, "_Sex_", s)]] <- fit_lmm_extract(sub, y, s, "Sex")
  }
  
  for (ag in age_levels) {
    sub <- data1 %>% filter(age_group == ag)
    res_lmm[[paste0(y, "_Age_", ag)]] <- fit_lmm_extract(sub, y, ag, "Age")
  }
  
  # ---------------- GEE (P75) ----------------
  res_gee[[paste0(y, "_Overall_All")]] <- fit_gee_extract(data1, y, "Overall", "All")
  
  for (s in sex_levels) {
    sub <- data1 %>% filter(sex_str == s)
    res_gee[[paste0(y, "_Sex_", s)]] <- fit_gee_extract(sub, y, "Sex", s)
  }
  
  for (ag in age_levels) {
    sub <- data1 %>% filter(age_group == ag)
    res_gee[[paste0(y, "_Age_", ag)]] <- fit_gee_extract(sub, y, "Age", ag)
  }
}

final_lmm <- bind_rows(res_lmm)
final_gee <- bind_rows(res_gee)

# ------------------------------------------------------------
# 7) 导出结果
# ------------------------------------------------------------
out_lmm <- file.path(out_dir, "LMM_fulladj_results.xlsx")
out_gee <- file.path(out_dir, "GEE_P75_fulladj_results.xlsx")

write.xlsx(final_lmm, out_lmm, overwrite = TRUE)
write.xlsx(final_gee, out_gee, overwrite = TRUE)

message("Done.")
message("LMM saved to: ", out_lmm)
message("GEE saved to: ", out_gee)

# ------------------------------------------------------------
# 8) 可选 QC：查看协变量的可用性（全样本）
# ------------------------------------------------------------
cov_exist <- intersect(c("bmi_m_pre", covars_full), names(data1))
qc_levels <- sapply(data1[, cov_exist, drop = FALSE], function(x) length(unique(na.omit(x))))
print(qc_levels)

cat("\nCheck key variable types:\n")
print(str(data1$log_hhincpc))
print(str(data1$AGE_m))
print(str(data1$V29_m))
print(str(data1$A12_m))



# ============================================================
# Visualize significant results (LMM betas + GEE ORs)
# - Input: LMM_fulladj_results.xlsx, GEE_P75_fulladj_results.xlsx
# - Output: Forest plots (PNG + PDF)
# ============================================================

rm(list = ls())

suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(stringr)
  library(tidyr)
  library(ggplot2)
  library(forcats)
  library(scales)
})

# ------------------------------------------------------------
# 0) 路径：与你导出结果保持一致
# ------------------------------------------------------------
out_dir <- "E:/博士课题/CHNS/CHNS_2013/初次整理/最终使用数据/食物变量再次处理/个人食物量表全匹配/FULLADJ_LMM_GEE_outputs"

lmm_xlsx <- file.path(out_dir, "LMM_fulladj_results.xlsx")
gee_xlsx <- file.path(out_dir, "GEE_P75_fulladj_results.xlsx")

stopifnot(file.exists(lmm_xlsx), file.exists(gee_xlsx))

# ------------------------------------------------------------
# 1) 读取结果
# ------------------------------------------------------------
lmm <- readxl::read_excel(lmm_xlsx)
gee <- readxl::read_excel(gee_xlsx)

# ------------------------------------------------------------
# 2) 统一处理：只保留显著结果
# ------------------------------------------------------------
alpha <- 0.05

lmm_sig <- lmm %>%
  filter(!is.na(p.value), p.value < alpha) %>%
  mutate(
    model = "LMM (continuous)",
    # 用于画图展示的标签
    effect_label = sprintf("%.2f (%.2f, %.2f)", estimate, conf.low, conf.high),
    stratum_label = paste0(stratum_type, ": ", stratum),
    outcome = as.character(outcome)
  )

gee_sig <- gee %>%
  filter(!is.na(p.value), p.value < alpha) %>%
  mutate(
    model = "GEE (P75 binary)",
    effect_label = sprintf("%.2f (%.2f, %.2f)", estimate, conf.low, conf.high),
    stratum_label = paste0(stratum_type, ": ", stratum),
    outcome = as.character(outcome)
  )

# 若没有显著结果，直接提示
if (nrow(lmm_sig) == 0) message("No significant LMM results at p<0.05.")
if (nrow(gee_sig) == 0) message("No significant GEE results at p<0.05.")

# ------------------------------------------------------------
# 3) 连续结局（LMM）Forest plot：β + 95%CI
#    - Facet: outcome
#    - y: strata (stratum_label)
# ------------------------------------------------------------
if (nrow(lmm_sig) > 0) {
  
  # 为了图形更美观：每个 outcome 内按效应大小排序
  lmm_sig2 <- lmm_sig %>%
    group_by(outcome) %>%
    mutate(stratum_label = fct_reorder(stratum_label, estimate)) %>%
    ungroup()
  
  p_lmm <- ggplot(lmm_sig2, aes(x = estimate, y = stratum_label)) +
    geom_vline(xintercept = 0, linewidth = 0.4) +
    geom_errorbarh(aes(xmin = conf.low, xmax = conf.high), height = 0.2) +
    geom_point(size = 2) +
    facet_wrap(~ outcome, scales = "free_y", ncol = 1) +
    labs(
      title = "Significant associations (LMM, continuous outcomes)",
      x = "Beta (95% CI) for bmi_m_pre",
      y = NULL
    ) +
    theme_bw(base_size = 11) +
    theme(
      strip.text = element_text(face = "bold"),
      panel.grid.minor = element_blank()
    )
  
  ggsave(file.path(out_dir, "LMM_significant_forest.png"), p_lmm, width = 9, height = 0.8 * max(6, nrow(lmm_sig2)), dpi = 300)
  ggsave(file.path(out_dir, "LMM_significant_forest.pdf"), p_lmm, width = 9, height = 0.8 * max(6, nrow(lmm_sig2)))
}

# ------------------------------------------------------------
# 4) 二分类结局（GEE）Forest plot：OR + 95%CI（log轴）
#    - Facet: outcome
#    - y: strata
# ------------------------------------------------------------
if (nrow(gee_sig) > 0) {
  
  gee_sig2 <- gee_sig %>%
    group_by(outcome) %>%
    mutate(stratum_label = fct_reorder(stratum_label, estimate)) %>%
    ungroup()
  
  p_gee <- ggplot(gee_sig2, aes(x = estimate, y = stratum_label)) +
    geom_vline(xintercept = 1, linewidth = 0.4) +
    geom_errorbarh(aes(xmin = conf.low, xmax = conf.high), height = 0.2) +
    geom_point(size = 2) +
    scale_x_log10(labels = scales::number_format(accuracy = 0.01)) +
    facet_wrap(~ outcome, scales = "free_y", ncol = 1) +
    labs(
      title = "Significant associations (GEE, P75 binary outcomes)",
      x = "Odds Ratio (95% CI) for bmi_m_pre (log scale)",
      y = NULL
    ) +
    theme_bw(base_size = 11) +
    theme(
      strip.text = element_text(face = "bold"),
      panel.grid.minor = element_blank()
    )
  
  ggsave(file.path(out_dir, "GEE_significant_forest.png"), p_gee, width = 9, height = 0.8 * max(6, nrow(gee_sig2)), dpi = 300)
  ggsave(file.path(out_dir, "GEE_significant_forest.pdf"), p_gee, width = 9, height = 0.8 * max(6, nrow(gee_sig2)))
}

# ------------------------------------------------------------
# 5)（可选）汇总一张“显著结果清单”便于核对
# ------------------------------------------------------------
sig_summary <- bind_rows(
  lmm_sig %>% transmute(model, outcome, stratum_type, stratum, estimate, conf.low, conf.high, p.value, covars_used),
  gee_sig %>% transmute(model, outcome, stratum_type, stratum, estimate, conf.low, conf.high, p.value, covars_used)
)

write.csv(sig_summary, file.path(out_dir, "significant_results_summary.csv"), row.names = FALSE)

message("Done. Plots and summary saved to: ", out_dir)



###############################################################
## CHNS Step1 (ONE-PASTE, RUNNABLE) + HORIZONTAL FACET PLOTS FIXED
## Requirements you specified:
## - Exposure: bmi_m_pre ONLY
## - Covariates: covars_full ONLY
## - Sex strata: All / Male (SEX_2=1) / Female (SEX_2=2)
## - Age group variable created: 0–6 / 7–10 / >10 (not used in models here)
## - Outcomes: WHO Z (LMM), WHO BIN + China BIN (GLMM)
## - PLOTS: ONLY horizontal facet forest plots (fixed axis + no missing outcome)
## Outputs:
## - Excel: 01_step1_growth/Step1_BMIpre_to_Growth_STRAT_SEX_fullcov.xlsx
## - Figures (fixed): 01_step1_growth/fig_strat_sex/
##   * Forest_CHINABIN_GLMM_bmi_m_pre_facetSex_FIXED.png
##   * Forest_WHOZ_LMM_bmi_m_pre_facetSex_FIXED.png
###############################################################

rm(list = ls())
options(dplyr.summarise.inform = FALSE)

suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(stringr)
  library(tidyr)
  library(purrr)
  library(lme4)
  library(lmerTest)
  library(broom.mixed)
  library(writexl)
  library(ggplot2)
})

###############################################################
## 0) PATHS (EDIT ONLY THESE)
###############################################################
path_data <- "E:/博士课题/CHNS/CHNS_2013/初次整理/最终使用数据/食物变量再次处理/个人食物量表全匹配/finaldata.xlsx"
dir_base  <- "E:/博士课题/CHNS/CHNS_2013/初次整理/最终使用数据/食物变量再次处理/个人食物量表全匹配/回归分析_finaldata_singleTable"

dir_step1 <- file.path(dir_base, "01_step1_growth")
dir_fig   <- file.path(dir_step1, "fig_strat_sex")
dir.create(dir_step1, recursive = TRUE, showWarnings = FALSE)
dir.create(dir_fig,   recursive = TRUE, showWarnings = FALSE)

###############################################################
## 1) CLEANING UTILS
###############################################################
clean_num <- function(x) {
  if (is.null(x)) return(NA_real_)
  if (is.numeric(x)) return(x)
  x_chr <- trimws(as.character(x))
  bad <- c("NA","na","N/A","n/a",""," ",".","-","--","NULL")
  x_chr[x_chr %in% bad] <- NA
  x_chr[grepl("[^0-9\\.\\-]", x_chr)] <- NA
  suppressWarnings(as.numeric(x_chr))
}

clean_int01 <- function(x) {
  v <- clean_num(x)
  v <- as.integer(v)
  v[!v %in% c(0L, 1L)] <- NA_integer_
  v
}

handle_income <- function(df) {
  if ("hhincpc_num" %in% names(df)) {
    hh <- clean_num(df$hhincpc_num)
  } else if ("hhincpc_cpi" %in% names(df)) {
    hh <- clean_num(df$hhincpc_cpi)
  } else {
    return(df)
  }
  df %>% mutate(log_hhincpc = ifelse(is.na(hh) | hh <= 0, NA_real_, log(hh)))
}

###############################################################
## 2) READ + PREP (types + age_group + sex_str)
###############################################################
cat("\n[00] Reading Excel...\n")
data1 <- read_excel(path_data, sheet = 1)

stopifnot(all(c("IDind","wave","SEX_2","age") %in% names(data1)))

data1 <- handle_income(data1)

# numeric coercion (if exist)
num_candidates <- intersect(c("age","AGE_m","hhsize","bmi_m_pre",
                              "haz","waz","whz","baz",
                              "WEIGHT_c","HEIGHT_c"), names(data1))
for (v in num_candidates) data1[[v]] <- clean_num(data1[[v]])

# bin outcomes coercion (if exist)
bin_candidates <- intersect(c(
  "haz_stunt_bin","waz_under_bin","baz_waste_bin","baz_over_bin","whz_waste_bin","whz_over_bin",
  "underweight_bin","stunting_bin","wasting_bin","owob_bin","obese_bin"
), names(data1))
for (v in bin_candidates) data1[[v]] <- clean_int01(data1[[v]])

# ----------------------------
# Covariates (exact list you gave)
# ----------------------------
covars_full <- c(
  "wave",
  "SEX_2",
  "han",
  "age",
  "A5C",
  "AGE_m",
  "V29_m",
  "A12_m",
  "hhsize",
  "region_ns_num",
  "index_tertile_num",
  "log_hhincpc"
)

# ----------------------------
# Create age_group and sex_str; factorize relevant vars
# ----------------------------
data1 <- data1 %>%
  mutate(
    age = as.numeric(age),
    age_group = case_when(
      !is.na(age) & age <= 6              ~ "0-6",
      !is.na(age) & age >= 7 & age <= 10  ~ "7-10",
      !is.na(age) & age > 10              ~ ">10",
      TRUE ~ NA_character_
    ),
    sex_str = case_when(
      SEX_2 == 1 ~ "Male",
      SEX_2 == 2 ~ "Female",
      TRUE ~ NA_character_
    ),
    IDind = as.factor(IDind)
  )

factor_vars <- intersect(
  c("wave","SEX_2","han","A5C","V29_m","A12_m","region_ns_num","index_tertile_num"),
  names(data1)
)
data1 <- data1 %>% mutate(across(all_of(factor_vars), as.factor))

###############################################################
## 3) EXPOSURE + OUTCOMES
###############################################################
exposure_main <- "bmi_m_pre"
stopifnot(exposure_main %in% names(data1))

outcomes_who_z     <- intersect(c("haz","waz","whz","baz"), names(data1))
outcomes_who_bin   <- intersect(c("haz_stunt_bin","waz_under_bin","baz_waste_bin","baz_over_bin","whz_waste_bin","whz_over_bin"), names(data1))
outcomes_china_bin <- intersect(c("underweight_bin","stunting_bin","wasting_bin","owob_bin","obese_bin"), names(data1))

cat("\n[01] Exposure:", exposure_main, "\n")
cat("[01] Covars  :", paste(intersect(covars_full, names(data1)), collapse=", "), "\n")
cat("[01] WHO Z   :", paste(outcomes_who_z, collapse=", "), "\n")
cat("[01] WHO BIN :", paste(outcomes_who_bin, collapse=", "), "\n")
cat("[01] CHN BIN :", paste(outcomes_china_bin, collapse=", "), "\n")

###############################################################
## 4) STRATA (All / Male / Female)
###############################################################
get_strata_list_sex <- function() {
  list(
    list(stratum_type="All", stratum="All",   filter_expr=NULL),
    list(stratum_type="Sex", stratum="Male",  filter_expr=quote(SEX_2 == 1)),
    list(stratum_type="Sex", stratum="Female",filter_expr=quote(SEX_2 == 2))
  )
}

# In Sex-stratified models, SEX_2 must be removed
adjust_covars_by_stratum <- function(covars, df_sub, stratum_type) {
  cov_use <- intersect(covars, names(df_sub))
  if (stratum_type == "Sex") cov_use <- setdiff(cov_use, "SEX_2")
  cov_use
}

# Drop 1-level factor covariates in the current subset (prevents contrasts error)
drop_one_level_factors <- function(d, covars) {
  cov_use <- intersect(covars, names(d))
  if (length(cov_use) == 0) return(character(0))
  keep <- c()
  for (v in cov_use) {
    x <- d[[v]]
    if (is.factor(x) || is.character(x)) {
      x <- droplevels(as.factor(x))
      if (nlevels(x) >= 2) keep <- c(keep, v)
    } else {
      if (!all(is.na(x))) keep <- c(keep, v)
    }
  }
  keep
}

###############################################################
## 5) FITTERS (safe)
###############################################################
fit_lmm <- function(df, outcome, exposure, covars, min_n = 80) {
  vars <- intersect(c(outcome, exposure, covars, "IDind"), names(df))
  d <- df %>% select(all_of(vars)) %>% drop_na()
  
  n <- nrow(d)
  n_id <- dplyr::n_distinct(d$IDind)
  
  if (n < min_n || is.na(n_id) || n_id < 2 || n_id >= n || dplyr::n_distinct(d[[exposure]]) < 2) {
    return(tibble(
      outcome=outcome, exposure=exposure, n=n, n_id=n_id,
      estimate=NA_real_, std.error=NA_real_, statistic=NA_real_, p.value=NA_real_,
      conf.low=NA_real_, conf.high=NA_real_,
      covars_used=NA_character_, dropped_1lvl=NA_character_,
      ok=FALSE, reason="Too sparse / no variation", is_singular=NA
    ))
  }
  
  cov_use <- drop_one_level_factors(d, covars)
  dropped <- setdiff(intersect(covars, names(d)), cov_use)
  
  rhs <- paste(c(exposure, cov_use), collapse=" + ")
  fml <- as.formula(paste0("`", outcome, "` ~ ", rhs, " + (1|IDind)"))
  
  m <- tryCatch(
    lmer(fml, data=d, REML=FALSE,
         control=lmerControl(optimizer="bobyqa", optCtrl=list(maxfun=20000))),
    error=function(e) e
  )
  
  if (inherits(m, "error")) {
    return(tibble(
      outcome=outcome, exposure=exposure, n=n, n_id=n_id,
      estimate=NA_real_, std.error=NA_real_, statistic=NA_real_, p.value=NA_real_,
      conf.low=NA_real_, conf.high=NA_real_,
      covars_used=paste(cov_use, collapse=" + "),
      dropped_1lvl=paste(dropped, collapse=" + "),
      ok=FALSE, reason=paste0("lmer failed: ", conditionMessage(m)),
      is_singular=NA
    ))
  }
  
  out <- broom.mixed::tidy(m, effects="fixed", conf.int=TRUE) %>%
    filter(term == exposure) %>%
    transmute(
      outcome, exposure, n, n_id,
      estimate, std.error, statistic, p.value,
      conf.low, conf.high
    )
  
  if (nrow(out) == 0) {
    out <- tibble(
      outcome=outcome, exposure=exposure, n=n, n_id=n_id,
      estimate=NA_real_, std.error=NA_real_, statistic=NA_real_, p.value=NA_real_,
      conf.low=NA_real_, conf.high=NA_real_
    )
  }
  
  out %>%
    mutate(
      covars_used=paste(cov_use, collapse=" + "),
      dropped_1lvl=paste(dropped, collapse=" + "),
      ok=TRUE, reason="OK",
      is_singular = isSingular(m, tol=1e-4)
    )
}

fit_glmm <- function(df, outcome_bin, exposure, covars, min_n=250, min_event=20) {
  vars <- intersect(c(outcome_bin, exposure, covars, "IDind","age"), names(df))
  d <- df %>% select(all_of(vars)) %>% drop_na()
  
  # WHO WHZ bins apply to age < 5 only
  if (outcome_bin %in% c("whz_waste_bin","whz_over_bin") && "age" %in% names(d)) {
    d <- d %>% filter(age < 5)
  }
  
  n <- nrow(d)
  n_id <- dplyr::n_distinct(d$IDind)
  n_event <- if (n > 0) sum(d[[outcome_bin]] == 1) else NA_integer_
  n_none  <- if (n > 0) sum(d[[outcome_bin]] == 0) else NA_integer_
  
  if (n < min_n || is.na(n_event) || n_event < min_event || n_none < min_event ||
      is.na(n_id) || n_id < 2 || n_id >= n || dplyr::n_distinct(d[[exposure]]) < 2) {
    return(tibble(
      outcome_bin=outcome_bin, exposure=exposure, n=n, n_id=n_id, n_event=n_event,
      estimate=NA_real_, std.error=NA_real_, statistic=NA_real_, p.value=NA_real_,
      conf.low=NA_real_, conf.high=NA_real_,
      OR=NA_real_, OR_low=NA_real_, OR_high=NA_real_,
      covars_used=NA_character_, dropped_1lvl=NA_character_,
      ok=FALSE, reason="Too sparse / events too few / no variation", is_singular=NA
    ))
  }
  
  cov_use <- drop_one_level_factors(d, covars)
  dropped <- setdiff(intersect(covars, names(d)), cov_use)
  
  rhs <- paste(c(exposure, cov_use), collapse=" + ")
  fml <- as.formula(paste0(outcome_bin, " ~ ", rhs, " + (1|IDind)"))
  
  m <- tryCatch(
    glmer(
      fml, data=d, family=binomial, nAGQ=0,
      control=glmerControl(optimizer="bobyqa", optCtrl=list(maxfun=20000))
    ),
    error=function(e) e
  )
  
  if (inherits(m, "error")) {
    return(tibble(
      outcome_bin=outcome_bin, exposure=exposure, n=n, n_id=n_id, n_event=n_event,
      estimate=NA_real_, std.error=NA_real_, statistic=NA_real_, p.value=NA_real_,
      conf.low=NA_real_, conf.high=NA_real_,
      OR=NA_real_, OR_low=NA_real_, OR_high=NA_real_,
      covars_used=paste(cov_use, collapse=" + "),
      dropped_1lvl=paste(dropped, collapse=" + "),
      ok=FALSE, reason=paste0("glmer failed: ", conditionMessage(m)),
      is_singular=NA
    ))
  }
  
  out <- broom.mixed::tidy(m, effects="fixed", conf.int=TRUE) %>%
    filter(term == exposure) %>%
    transmute(
      outcome_bin, exposure, n, n_id, n_event,
      estimate, std.error, statistic, p.value,
      conf.low, conf.high,
      OR = exp(estimate),
      OR_low = exp(conf.low),
      OR_high = exp(conf.high)
    )
  
  if (nrow(out) == 0) {
    out <- tibble(
      outcome_bin=outcome_bin, exposure=exposure, n=n, n_id=n_id, n_event=n_event,
      estimate=NA_real_, std.error=NA_real_, statistic=NA_real_, p.value=NA_real_,
      conf.low=NA_real_, conf.high=NA_real_,
      OR=NA_real_, OR_low=NA_real_, OR_high=NA_real_
    )
  }
  
  out %>%
    mutate(
      covars_used=paste(cov_use, collapse=" + "),
      dropped_1lvl=paste(dropped, collapse=" + "),
      ok=TRUE, reason="OK",
      is_singular = isSingular(m, tol=1e-4)
    )
}

###############################################################
## 6) RUN MODELS (sex-stratified for ALL outcomes)
###############################################################
run_strata_lmm <- function(df, outcomes, exposure, covars) {
  strata <- get_strata_list_sex()
  map_dfr(strata, function(st) {
    dsub <- df
    if (!is.null(st$filter_expr)) dsub <- dsub %>% filter(!!st$filter_expr)
    
    cov_use0 <- adjust_covars_by_stratum(covars, dsub, st$stratum_type)
    
    map_dfr(outcomes, function(y) {
      fit_lmm(dsub, outcome=y, exposure=exposure, covars=cov_use0) %>%
        mutate(stratum_type=st$stratum_type, stratum=st$stratum)
    })
  }) %>%
    group_by(stratum_type, stratum) %>%
    mutate(p_fdr = p.adjust(p.value, "fdr")) %>%
    ungroup()
}

run_strata_glmm <- function(df, outcomes_bin, exposure, covars) {
  strata <- get_strata_list_sex()
  map_dfr(strata, function(st) {
    dsub <- df
    if (!is.null(st$filter_expr)) dsub <- dsub %>% filter(!!st$filter_expr)
    
    cov_use0 <- adjust_covars_by_stratum(covars, dsub, st$stratum_type)
    
    map_dfr(outcomes_bin, function(y) {
      fit_glmm(dsub, outcome_bin=y, exposure=exposure, covars=cov_use0) %>%
        mutate(stratum_type=st$stratum_type, stratum=st$stratum)
    })
  }) %>%
    group_by(stratum_type, stratum) %>%
    mutate(p_fdr = p.adjust(p.value, "fdr")) %>%
    ungroup()
}

cat("\n[02] Running Step1 models...\n")
covars_use_full <- intersect(covars_full, names(data1))

res_who_z     <- run_strata_lmm(data1, outcomes_who_z, exposure_main, covars_use_full)
res_who_bin   <- run_strata_glmm(data1, outcomes_who_bin, exposure_main, covars_use_full)
res_china_bin <- run_strata_glmm(data1, outcomes_china_bin, exposure_main, covars_use_full)

out_xlsx <- file.path(dir_step1, "Step1_BMIpre_to_Growth_STRAT_SEX_fullcov.xlsx")
writexl::write_xlsx(list(
  WHO_Z_LMM      = res_who_z,
  WHO_BIN_GLMM   = res_who_bin,
  CHINA_BIN_GLMM = res_china_bin
), out_xlsx)
cat("✅ Excel saved:", out_xlsx, "\n")

###############################################################
## 7) PLOTS (ONLY HORIZONTAL FACETS, FIXED AXIS + COMPLETE)
###############################################################
facet_levels <- c("All","Male","Female")
facet_label <- function(stratum_type, stratum){
  ifelse(stratum_type == "All", "All",
         ifelse(stratum %in% c("Male","M"), "Male",
                ifelse(stratum %in% c("Female","F"), "Female",
                       paste0(stratum_type, "=", stratum))))
}
theme_forest_h <- function(base_size = 12){
  theme_bw(base_size = base_size) +
    theme(
      panel.grid.major.y = element_blank(),
      panel.grid.minor.y = element_blank(),
      strip.background = element_rect(fill = "grey85", color = "grey40"),
      strip.text = element_text(size = base_size),
      plot.title = element_text(hjust = 0.5, face = "bold"),
      plot.subtitle = element_text(hjust = 0.5),
      axis.title.y = element_text(margin = margin(r = 6)),
      axis.title.x = element_text(margin = margin(t = 6))
    )
}

plot_or_facet_h <- function(res,
                            title_main,
                            subtitle = NULL,
                            or_limits = c(0.6, 1.6),
                            or_breaks = c(0.6, 0.8, 1.0, 1.25, 1.6),
                            base_size = 12){
  
  stopifnot(all(c("outcome_bin","OR","OR_low","OR_high","n","n_event","stratum_type","stratum") %in% names(res)))
  
  df <- res %>%
    mutate(
      facet = factor(facet_label(stratum_type, stratum), levels = facet_levels),
      outcome_bin = as.character(outcome_bin)
    )
  
  outcome_order <- c("underweight_bin","stunting_bin","wasting_bin","owob_bin","obese_bin",
                     "haz_stunt_bin","waz_under_bin","baz_waste_bin","baz_over_bin","whz_waste_bin","whz_over_bin")
  outcome_levels <- intersect(outcome_order, unique(df$outcome_bin))
  if (length(outcome_levels) == 0) outcome_levels <- unique(df$outcome_bin)
  
  df2 <- df %>%
    mutate(outcome_bin = factor(outcome_bin, levels = outcome_levels)) %>%
    select(facet, outcome_bin, n, n_event, OR, OR_low, OR_high) %>%
    complete(
      facet = factor(facet_levels, levels = facet_levels),
      outcome_bin = factor(outcome_levels, levels = outcome_levels)
    ) %>%
    mutate(
      label = ifelse(is.na(n),
                     as.character(outcome_bin),
                     paste0(as.character(outcome_bin), " (n=", n, ", event=", n_event, ")")),
      label = factor(label, levels = rev(unique(label)))
    )
  
  ggplot(df2, aes(x = label, y = OR)) +
    geom_hline(yintercept = 1, linetype = "dashed") +
    geom_pointrange(aes(ymin = OR_low, ymax = OR_high),
                    linewidth = 0.6, na.rm = TRUE) +
    scale_y_log10(limits = or_limits, breaks = or_breaks) +
    coord_flip() +
    facet_wrap(~facet, nrow = 1, scales = "free_y") +
    labs(x = "结局", y = "OR (95% CI)", title = title_main, subtitle = subtitle) +
    theme_forest_h(base_size = base_size)
}

plot_beta_facet_h <- function(res,
                              title_main,
                              subtitle = NULL,
                              beta_limits = c(-1.0, 1.0),
                              beta_breaks = seq(-1.0, 1.0, by = 0.5),
                              base_size = 12){
  
  stopifnot(all(c("outcome","estimate","conf.low","conf.high","n","stratum_type","stratum") %in% names(res)))
  
  df <- res %>%
    mutate(
      facet = factor(facet_label(stratum_type, stratum), levels = facet_levels),
      outcome = as.character(outcome)
    )
  
  outcome_levels <- intersect(c("haz","waz","whz","baz"), unique(df$outcome))
  if (length(outcome_levels) == 0) outcome_levels <- unique(df$outcome)
  
  df2 <- df %>%
    mutate(outcome = factor(outcome, levels = outcome_levels)) %>%
    select(facet, outcome, n, estimate, conf.low, conf.high) %>%
    complete(
      facet = factor(facet_levels, levels = facet_levels),
      outcome = factor(outcome_levels, levels = outcome_levels)
    ) %>%
    mutate(
      label = ifelse(is.na(n),
                     as.character(outcome),
                     paste0(as.character(outcome), " (n=", n, ")")),
      label = factor(label, levels = rev(unique(label)))
    )
  
  ggplot(df2, aes(x = label, y = estimate)) +
    geom_hline(yintercept = 0, linetype = "dashed") +
    geom_pointrange(aes(ymin = conf.low, ymax = conf.high),
                    linewidth = 0.6, na.rm = TRUE) +
    scale_y_continuous(limits = beta_limits, breaks = beta_breaks) +
    coord_flip() +
    facet_wrap(~facet, nrow = 1, scales = "free_y") +
    labs(x = "WHO Z 指标", y = "β (95% CI)", title = title_main, subtitle = subtitle) +
    theme_forest_h(base_size = base_size)
}

cat("\n[05] Plotting (horizontal facets only)...\n")

# China BIN (GLMM) fixed plot
p_china <- plot_or_facet_h(
  res = res_china_bin,
  title_main = "孕前母亲 BMI（bmi_m_pre）与子代体格异常（中国标准，GLMM）",
  subtitle   = "全协变量调整；分层：All / Male / Female；随机截距：(1|IDind)",
  or_limits  = c(0.6, 1.6),
  or_breaks  = c(0.6, 0.8, 1.0, 1.25, 1.6)
)
out_png_china <- file.path(dir_fig, "Forest_CHINABIN_GLMM_bmi_m_pre_facetSex_FIXED.png")
ggsave(out_png_china, p_china, width = 11.6, height = 6.0, dpi = 300)
cat("✅ Saved:", out_png_china, "\n")

# ============================================================
# WHO Z (LMM) — HORIZONTAL FACET FOREST (NARROW X-AXIS)
# ============================================================

p_whoz <- plot_beta_facet_h(
  res = res_who_z,
  title_main = "孕前母亲 BMI（bmi_m_pre）与子代 WHO Z（LMM）",
  subtitle   = "全协变量调整；分层：All / Male / Female；随机截距：(1|IDind)",
  
  # 🔑 关键修改在这里
  beta_limits = c(-0.5, 0.5),
  beta_breaks = seq(-0.5, 0.5, by = 0.25)
)

ggsave(
  "Forest_WHOZ_LMM_bmi_m_pre_facetSex_NARROW.png",
  p_whoz,
  width  = 11.2,
  height = 5.8,
  dpi    = 300
)




# 按照菌群分析流程 去处理食物！！！重新分析
###############################################################
## CHNS Dietary Pipeline (Personal food-entry -> daily -> wave mean3
## -> density vars -> cumulative average) + Excel multi-sheets
## ✅ Food category summarized as ENERGY (kcal/day), NOT grams.
##
## Input:
## - Excel from your FCT-matching step: data_food_matched
##   Must contain: IDind, wave, vd, V39, edible_pct, food_category,
##   nutrient density cols (per 100g edible): energy_kcal, protein_g, fat_g, ...
##
## Output (one Excel with multi-sheets):
## - 00_food_entry_std
## - 01_daily_sum
## - 02_wave_3day_mean
## - 03_wave_density
## - 04_cumavg
## - 05_analysis_ready_diet
###############################################################

rm(list = ls())

suppressPackageStartupMessages({
  library(dplyr)
  library(tidyr)
  library(stringr)
  library(readxl)
  library(openxlsx)
})

options(dplyr.summarise.inform = FALSE)

###############################################################
## 0) Paths (EDIT THESE)
###############################################################
in_path  <- "E:/博士课题/CHNS/CHNS_2013/初次整理/最终使用数据/食物变量再次处理/个人食物量表全匹配/个人食物条目_仅匹配FCT_含食物大类_输出与QC.xlsx"
in_sheet <- 1
out_xlsx <- "E:/博士课题/CHNS/CHNS_2013/初次整理/最终使用数据/食物变量再次处理/个人食物量表全匹配饮食暴露_全流程_daily_wave_cumavg_KCAL_BY_CAT.xlsx"

###############################################################
## 1) Helper functions
###############################################################
clean_na_char <- function(x) {
  if (is.numeric(x)) return(x)
  x_chr <- trimws(as.character(x))
  bad <- c("NA","na","N/A","n/a",""," ",".","..","-","--","NULL","null","NaN","nan")
  x_chr[x_chr %in% bad] <- NA_character_
  x_chr
}

to_num <- function(x) suppressWarnings(as.numeric(clean_na_char(x)))

# Safe divide
safe_div <- function(a, b) ifelse(is.na(b) | b == 0, NA_real_, a / b)

# Cumulative mean for numeric columns within IDind ordered by wave
cummean_vec <- function(x) {
  out <- rep(NA_real_, length(x))
  s <- 0
  n <- 0
  for (i in seq_along(x)) {
    if (!is.na(x[i])) {
      s <- s + x[i]
      n <- n + 1
    }
    out[i] <- ifelse(n == 0, NA_real_, s / n)
  }
  out
}

###############################################################
## 2) Read data (matched food entries)
###############################################################
df0 <- read_excel(in_path, sheet = in_sheet) %>%
  mutate(
    IDind = as.character(IDind),
    wave  = as.character(wave),
    vd    = as.character(vd)
  )

stopifnot(nrow(df0) > 0)

###############################################################
## 3) Define nutrient density fields (per 100g edible)
###############################################################
nutr_cols <- c(
  "energy_kj","energy_kcal",
  "water_g","protein_g","fat_g","fiber_total_g","carb_g","ash_g",
  "vitA_ugRE","carotene_ug",
  "thiamin_mg","riboflavin_mg","niacin_mg",
  "vitC_mg","vitE_total_mg","vitE_alpha_mg",
  "k_mg","na_mg","ca_mg","mg_mg","fe_mg","mn_mg","zn_mg","cu_mg","p_mg","se_ug"
)

missing_nutr <- setdiff(nutr_cols, names(df0))
if (length(missing_nutr) > 0) {
  stop("Missing nutrient columns in input: ", paste(missing_nutr, collapse = ", "))
}

req_cols <- c("IDind","vd","wave","V39","edible_pct","food_category")
missing_req <- setdiff(req_cols, names(df0))
if (length(missing_req) > 0) {
  stop("Missing required columns in input: ", paste(missing_req, collapse = ", "))
}

###############################################################
## 4) Sheet: 00_food_entry_std
##    - Harmonize V39 into grams (food_g_std)
##    - Compute edible grams (edible_g)
##    - Compute per-entry nutrient intakes for ALL nutrient cols
###############################################################
df_entry <- df0 %>%
  mutate(
    wave_num       = suppressWarnings(as.integer(clean_na_char(wave))),
    V39_num        = to_num(V39),
    edible_pct_num = to_num(edible_pct),
    
    # Unit harmonization: 1997-2000 are liang; >=2004 grams
    food_g_std = case_when(
      !is.na(wave_num) & wave_num >= 1997 & wave_num <= 2000 ~ V39_num * 50,
      !is.na(wave_num) & wave_num >= 2004 ~ V39_num,
      TRUE ~ V39_num  # fallback
    ),
    edible_g = food_g_std * edible_pct_num / 100
  ) %>%
  mutate(across(all_of(nutr_cols), to_num))

# Compute entry-level nutrient intakes:
# Assumption: nutrient density columns are per 100g edible portion.
for (cc in nutr_cols) {
  new_nm <- paste0(cc, "_intake")
  df_entry[[new_nm]] <- (df_entry$edible_g / 100) * df_entry[[cc]]
}

# Safety check (this is required for kcal-by-category)
stopifnot("energy_kcal_intake" %in% names(df_entry))

###############################################################
## 5) Sheet: 01_daily_sum (IDind x wave x vd)
##    - Sum nutrient intakes per day
##    - Sum edible grams per day (optional QC)
##    - Sum kcal by food_category per day (wide)   ✅ key change
###############################################################
intake_cols <- paste0(nutr_cols, "_intake")

# Daily nutrient totals
df_daily_nutr <- df_entry %>%
  group_by(IDind, wave, vd) %>%
  summarise(
    n_food_items_day = dplyr::n(),
    total_food_g_day = sum(edible_g, na.rm = TRUE),
    across(all_of(intake_cols), ~ sum(.x, na.rm = TRUE)),
    .groups = "drop"
  )

# Rename nutrient intake columns to *_day (remove _intake suffix)
names(df_daily_nutr) <- names(df_daily_nutr) %>%
  str_replace("_intake$", "_day")

# ✅ Daily food category kcal (wide)
df_daily_cat <- df_entry %>%
  mutate(
    food_category = clean_na_char(food_category),
    food_category = trimws(food_category)
  ) %>%
  filter(!is.na(food_category), food_category != "") %>%
  group_by(IDind, wave, vd, food_category) %>%
  summarise(
    cat_kcal_day = sum(energy_kcal_intake, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  mutate(food_category = make.names(food_category)) %>%
  pivot_wider(
    names_from  = food_category,
    values_from = cat_kcal_day,
    values_fill = 0
  ) %>%
  rename_with(~ paste0(.x, "_kcal_day"), -c(IDind, wave, vd))

# Combine daily nutrients + category kcal
df_daily <- df_daily_nutr %>%
  left_join(df_daily_cat, by = c("IDind","wave","vd"))

###############################################################
## 6) Sheet: 02_wave_3day_mean (IDind x wave)
##    - Mean across vd within wave
##    - Keep n_days_recorded, total items, etc.
###############################################################
daily_num_cols <- df_daily %>%
  select(-IDind, -wave, -vd) %>%
  names()

df_wave_mean3 <- df_daily %>%
  group_by(IDind, wave) %>%
  summarise(
    n_days_recorded = dplyr::n_distinct(vd),
    n_items_wave    = sum(n_food_items_day, na.rm = TRUE),
    across(all_of(daily_num_cols), ~ mean(.x, na.rm = TRUE)),
    .groups = "drop"
  )

# Rename averaged daily cols: *_day -> *_mean3
names(df_wave_mean3) <- names(df_wave_mean3) %>%
  str_replace("_day$", "_mean3")

###############################################################
## 7) Sheet: 03_wave_density
##    - Derive energy shares and density variables from mean3
###############################################################
must_exist <- c("energy_kcal_mean3","protein_g_mean3","fat_g_mean3","carb_g_mean3",
                "fiber_total_g_mean3","total_food_g_mean3")
stopifnot(all(must_exist %in% names(df_wave_mean3)))

df_wave_density <- df_wave_mean3 %>%
  transmute(
    IDind, wave,
    energy_kcal_mean3,
    protein_g_mean3,
    fat_g_mean3,
    carb_g_mean3,
    fiber_total_g_mean3,
    total_food_g_mean3,
    fat_pct_energy        = safe_div(fat_g_mean3 * 9,  energy_kcal_mean3) * 100,
    protein_pct_energy    = safe_div(protein_g_mean3 * 4, energy_kcal_mean3) * 100,
    carb_pct_energy       = safe_div(carb_g_mean3 * 4, energy_kcal_mean3) * 100,
    fiber_g_per_1kcal     = safe_div(fiber_total_g_mean3, energy_kcal_mean3),
    fiber_g_per_1000kcal  = safe_div(fiber_total_g_mean3, energy_kcal_mean3) * 1000,
    energy_density_kcal_per_g = safe_div(energy_kcal_mean3, total_food_g_mean3)
  )

###############################################################
## 8) Sheet: 04_cumavg (IDind x wave)
##    - Cumulative averages for ALL wave-level dietary variables
###############################################################
df_wave_all <- df_wave_mean3 %>%
  left_join(
    df_wave_density %>%
      select(-energy_kcal_mean3, -protein_g_mean3, -fat_g_mean3, -carb_g_mean3,
             -fiber_total_g_mean3, -total_food_g_mean3),
    by = c("IDind","wave")
  ) %>%
  mutate(wave_num = suppressWarnings(as.integer(clean_na_char(wave)))) %>%
  arrange(IDind, wave_num, wave)

id_cols <- c("IDind","wave","wave_num")
diet_num_cols <- df_wave_all %>%
  select(-all_of(id_cols)) %>%
  select(where(is.numeric)) %>%
  names()

df_cumavg <- df_wave_all %>%
  group_by(IDind) %>%
  arrange(wave_num, .by_group = TRUE) %>%
  mutate(
    across(all_of(diet_num_cols), ~ cummean_vec(.x), .names = "cum_{.col}")
  ) %>%
  ungroup() %>%
  select(IDind, wave, starts_with("cum_"))

###############################################################
## 9) Sheet: 05_analysis_ready_diet
##    - Final analysis table: mean3 + density + cumavg
###############################################################
df_analysis_ready <- df_wave_all %>%
  select(-wave_num) %>%
  left_join(df_cumavg, by = c("IDind","wave"))

###############################################################
## 10) Write all sheets to one Excel
###############################################################
wb <- createWorkbook()

addWorksheet(wb, "00_food_entry_std")
writeData(wb, "00_food_entry_std", df_entry)

addWorksheet(wb, "01_daily_sum")
writeData(wb, "01_daily_sum", df_daily)

addWorksheet(wb, "02_wave_3day_mean")
writeData(wb, "02_wave_3day_mean", df_wave_mean3)

addWorksheet(wb, "03_wave_density")
writeData(wb, "03_wave_density", df_wave_density)

addWorksheet(wb, "04_cumavg")
writeData(wb, "04_cumavg", df_cumavg)

addWorksheet(wb, "05_analysis_ready_diet")
writeData(wb, "05_analysis_ready_diet", df_analysis_ready)

saveWorkbook(wb, out_xlsx, overwrite = TRUE)

message("✅ Finished. Output saved to: ", out_xlsx)

# Optional QC prints
cat("\n[QC] df_daily_cat: rows=", nrow(df_daily_cat),
    " | kcal-category cols=", ncol(df_daily_cat) - 3, "\n")
print(head(df_daily_cat, 3))
cat("\n[QC] df_wave_mean3: rows=", nrow(df_wave_mean3),
    " | cols=", ncol(df_wave_mean3), "\n")



# 将含有没小类食物的kcal合并至含有基线变量的表格中

library(readxl)
library(dplyr)
library(openxlsx)

# 读入数据
data1 <- read_excel(
  "E:/博士课题/CHNS/CHNS_2013/初次整理/最终使用数据/食物变量再次处理/个人食物量表全匹配/调整城市化指数变量.xlsx",
  sheet = 1
)

data2 <- read_excel(
  "E:/博士课题/CHNS/CHNS_2013/初次整理/最终使用数据/食物变量再次处理/个人食物量表全匹配/个人食物量表全匹配饮食暴露_全流程_daily_wave_cumavg_KCAL_BY_CAT.xlsx",
  sheet = 6
)

# 统一 join key 类型
data1 <- data1 %>%
  mutate(
    IDind = as.character(IDind),
    wave  = as.character(wave)
  )

data2 <- data2 %>%
  mutate(
    IDind = as.character(IDind),
    wave  = as.character(wave)
  )

# 正确合并（推荐）
data_merged1 <- data1 %>%
  left_join(data2, by = c("IDind", "wave"))

# 基本检查
stopifnot(nrow(data_merged1) == nrow(data1))
cat("合并后行数：", nrow(data_merged1), "\n")

# 导出
write.xlsx(
  data_merged1,
  "E:/博士课题/CHNS/CHNS_2013/初次整理/最终使用数据/食物变量再次处理/个人食物量表全匹配/finaldata_kcal.xlsx",
  overwrite = TRUE
)



