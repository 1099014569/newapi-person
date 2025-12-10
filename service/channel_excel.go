package service

import (
	"encoding/json"
	"fmt"
	"io"
	"strconv"

	"github.com/QuantumNous/new-api/model"
	"github.com/xuri/excelize/v2"
)

// ChannelExcelParseResult Excel 解析结果
type ChannelExcelParseResult struct {
	Channels []model.Channel
	Errors   []string
}

// GenerateChannelsExcel 生成渠道 Excel 文件
func GenerateChannelsExcel(channels []model.Channel) (*excelize.File, error) {
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	sheetName := "Channels"
	index, err := f.NewSheet(sheetName)
	if err != nil {
		return nil, err
	}

	// 删除默认的 Sheet1
	f.DeleteSheet("Sheet1")

	// 设置表头
	headers := []string{
		"ID", "类型", "名称", "密钥", "状态", "分组", "模型列表", "BaseURL",
		"优先级", "权重", "标签", "备注", "余额", "已用配额", "响应时间",
		"测试模型", "OpenAI组织", "其他配置", "模型映射", "状态码映射",
		"自动封禁", "渠道设置", "参数覆盖", "请求头覆盖", "其他设置",
		"创建时间", "测试时间", "余额更新时间", "渠道信息JSON",
	}

	// 写入表头
	for i, header := range headers {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		f.SetCellValue(sheetName, cell, header)
	}

	// 设置表头样式
	style, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#4472C4"},
			Pattern: 1,
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	})
	f.SetCellStyle(sheetName, "A1", "AC1", style)

	// 写入数据
	for i, channel := range channels {
		row := i + 2

		// 序列化 JSON 字段
		channelInfoJSON, _ := json.Marshal(channel.ChannelInfo)

		// 处理指针类型字段
		baseURL := ""
		if channel.BaseURL != nil {
			baseURL = *channel.BaseURL
		}
		priority := int64(0)
		if channel.Priority != nil {
			priority = *channel.Priority
		}
		weight := uint(0)
		if channel.Weight != nil {
			weight = *channel.Weight
		}
		tag := ""
		if channel.Tag != nil {
			tag = *channel.Tag
		}
		remark := ""
		if channel.Remark != nil {
			remark = *channel.Remark
		}
		testModel := ""
		if channel.TestModel != nil {
			testModel = *channel.TestModel
		}
		openaiOrg := ""
		if channel.OpenAIOrganization != nil {
			openaiOrg = *channel.OpenAIOrganization
		}
		modelMapping := ""
		if channel.ModelMapping != nil {
			modelMapping = *channel.ModelMapping
		}
		statusCodeMapping := ""
		if channel.StatusCodeMapping != nil {
			statusCodeMapping = *channel.StatusCodeMapping
		}
		autoBan := 1
		if channel.AutoBan != nil {
			autoBan = *channel.AutoBan
		}
		setting := ""
		if channel.Setting != nil {
			setting = *channel.Setting
		}
		paramOverride := ""
		if channel.ParamOverride != nil {
			paramOverride = *channel.ParamOverride
		}
		headerOverride := ""
		if channel.HeaderOverride != nil {
			headerOverride = *channel.HeaderOverride
		}

		// 写入每列数据
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), channel.Id)
		f.SetCellValue(sheetName, fmt.Sprintf("B%d", row), channel.Type)
		f.SetCellValue(sheetName, fmt.Sprintf("C%d", row), channel.Name)
		f.SetCellValue(sheetName, fmt.Sprintf("D%d", row), channel.Key)
		f.SetCellValue(sheetName, fmt.Sprintf("E%d", row), channel.Status)
		f.SetCellValue(sheetName, fmt.Sprintf("F%d", row), channel.Group)
		f.SetCellValue(sheetName, fmt.Sprintf("G%d", row), channel.Models)
		f.SetCellValue(sheetName, fmt.Sprintf("H%d", row), baseURL)
		f.SetCellValue(sheetName, fmt.Sprintf("I%d", row), priority)
		f.SetCellValue(sheetName, fmt.Sprintf("J%d", row), weight)
		f.SetCellValue(sheetName, fmt.Sprintf("K%d", row), tag)
		f.SetCellValue(sheetName, fmt.Sprintf("L%d", row), remark)
		f.SetCellValue(sheetName, fmt.Sprintf("M%d", row), channel.Balance)
		f.SetCellValue(sheetName, fmt.Sprintf("N%d", row), channel.UsedQuota)
		f.SetCellValue(sheetName, fmt.Sprintf("O%d", row), channel.ResponseTime)
		f.SetCellValue(sheetName, fmt.Sprintf("P%d", row), testModel)
		f.SetCellValue(sheetName, fmt.Sprintf("Q%d", row), openaiOrg)
		f.SetCellValue(sheetName, fmt.Sprintf("R%d", row), channel.Other)
		f.SetCellValue(sheetName, fmt.Sprintf("S%d", row), modelMapping)
		f.SetCellValue(sheetName, fmt.Sprintf("T%d", row), statusCodeMapping)
		f.SetCellValue(sheetName, fmt.Sprintf("U%d", row), autoBan)
		f.SetCellValue(sheetName, fmt.Sprintf("V%d", row), setting)
		f.SetCellValue(sheetName, fmt.Sprintf("W%d", row), paramOverride)
		f.SetCellValue(sheetName, fmt.Sprintf("X%d", row), headerOverride)
		f.SetCellValue(sheetName, fmt.Sprintf("Y%d", row), channel.OtherSettings)
		f.SetCellValue(sheetName, fmt.Sprintf("Z%d", row), channel.CreatedTime)
		f.SetCellValue(sheetName, fmt.Sprintf("AA%d", row), channel.TestTime)
		f.SetCellValue(sheetName, fmt.Sprintf("AB%d", row), channel.BalanceUpdatedTime)
		f.SetCellValue(sheetName, fmt.Sprintf("AC%d", row), string(channelInfoJSON))
	}

	// 设置列宽
	f.SetColWidth(sheetName, "A", "A", 10)
	f.SetColWidth(sheetName, "B", "B", 10)
	f.SetColWidth(sheetName, "C", "C", 20)
	f.SetColWidth(sheetName, "D", "D", 40)
	f.SetColWidth(sheetName, "E", "E", 10)
	f.SetColWidth(sheetName, "F", "F", 15)
	f.SetColWidth(sheetName, "G", "G", 40)
	f.SetColWidth(sheetName, "H", "H", 40)
	f.SetColWidth(sheetName, "I", "I", 10)
	f.SetColWidth(sheetName, "J", "J", 10)
	f.SetColWidth(sheetName, "K", "K", 15)
	f.SetColWidth(sheetName, "L", "L", 30)

	// 设置活动工作表
	f.SetActiveSheet(index)

	return f, nil
}

// ParseChannelsExcel 解析渠道 Excel 文件
func ParseChannelsExcel(reader io.Reader) (*ChannelExcelParseResult, error) {
	f, err := excelize.OpenReader(reader)
	if err != nil {
		return nil, err
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// 获取第一个 sheet
	sheetName := f.GetSheetName(0)
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, err
	}

	if len(rows) < 2 {
		return nil, fmt.Errorf("Excel 文件为空或格式不正确")
	}

	var channels []model.Channel
	var errors []string

	// 跳过表头，从第二行开始
	for i := 1; i < len(rows); i++ {
		row := rows[i]
		if len(row) == 0 {
			continue // 跳过空行
		}

		channel := model.Channel{}

		// 解析每一列（使用辅助函数处理越界）
		getCell := func(index int) string {
			if index < len(row) {
				return row[index]
			}
			return ""
		}

		// ID (可选，如果为空或已存在则忽略)
		if idStr := getCell(0); idStr != "" {
			if id, err := strconv.Atoi(idStr); err == nil {
				channel.Id = id
			}
		}

		// 类型 (必填)
		if typeStr := getCell(1); typeStr != "" {
			if channelType, err := strconv.Atoi(typeStr); err == nil {
				channel.Type = channelType
			}
		}

		// 名称 (必填)
		channel.Name = getCell(2)

		// 密钥 (必填)
		channel.Key = getCell(3)

		// 状态
		if statusStr := getCell(4); statusStr != "" {
			if status, err := strconv.Atoi(statusStr); err == nil {
				channel.Status = status
			} else {
				channel.Status = 1 // 默认启用
			}
		} else {
			channel.Status = 1
		}

		// 分组
		if group := getCell(5); group != "" {
			channel.Group = group
		} else {
			channel.Group = "default"
		}

		// 模型列表
		channel.Models = getCell(6)

		// BaseURL
		if baseURL := getCell(7); baseURL != "" {
			channel.BaseURL = &baseURL
		}

		// 优先级
		if priorityStr := getCell(8); priorityStr != "" {
			if priority, err := strconv.ParseInt(priorityStr, 10, 64); err == nil {
				channel.Priority = &priority
			}
		}

		// 权重
		if weightStr := getCell(9); weightStr != "" {
			if weight, err := strconv.ParseUint(weightStr, 10, 32); err == nil {
				w := uint(weight)
				channel.Weight = &w
			}
		}

		// 标签
		if tag := getCell(10); tag != "" {
			channel.Tag = &tag
		}

		// 备注
		if remark := getCell(11); remark != "" {
			channel.Remark = &remark
		}

		// 余额
		if balanceStr := getCell(12); balanceStr != "" {
			if balance, err := strconv.ParseFloat(balanceStr, 64); err == nil {
				channel.Balance = balance
			}
		}

		// 已用配额
		if usedQuotaStr := getCell(13); usedQuotaStr != "" {
			if usedQuota, err := strconv.ParseInt(usedQuotaStr, 10, 64); err == nil {
				channel.UsedQuota = usedQuota
			}
		}

		// 响应时间
		if responseTimeStr := getCell(14); responseTimeStr != "" {
			if responseTime, err := strconv.Atoi(responseTimeStr); err == nil {
				channel.ResponseTime = responseTime
			}
		}

		// 测试模型
		if testModel := getCell(15); testModel != "" {
			channel.TestModel = &testModel
		}

		// OpenAI 组织
		if openaiOrg := getCell(16); openaiOrg != "" {
			channel.OpenAIOrganization = &openaiOrg
		}

		// 其他配置
		channel.Other = getCell(17)

		// 模型映射
		if modelMapping := getCell(18); modelMapping != "" {
			channel.ModelMapping = &modelMapping
		}

		// 状态码映射
		if statusCodeMapping := getCell(19); statusCodeMapping != "" {
			channel.StatusCodeMapping = &statusCodeMapping
		}

		// 自动封禁
		if autoBanStr := getCell(20); autoBanStr != "" {
			if autoBan, err := strconv.Atoi(autoBanStr); err == nil {
				channel.AutoBan = &autoBan
			}
		}

		// 渠道设置
		if setting := getCell(21); setting != "" {
			channel.Setting = &setting
		}

		// 参数覆盖
		if paramOverride := getCell(22); paramOverride != "" {
			channel.ParamOverride = &paramOverride
		}

		// 请求头覆盖
		if headerOverride := getCell(23); headerOverride != "" {
			channel.HeaderOverride = &headerOverride
		}

		// 其他设置
		channel.OtherSettings = getCell(24)

		// 创建时间
		if createdTimeStr := getCell(25); createdTimeStr != "" {
			if createdTime, err := strconv.ParseInt(createdTimeStr, 10, 64); err == nil {
				channel.CreatedTime = createdTime
			}
		}

		// 测试时间
		if testTimeStr := getCell(26); testTimeStr != "" {
			if testTime, err := strconv.ParseInt(testTimeStr, 10, 64); err == nil {
				channel.TestTime = testTime
			}
		}

		// 余额更新时间
		if balanceUpdatedTimeStr := getCell(27); balanceUpdatedTimeStr != "" {
			if balanceUpdatedTime, err := strconv.ParseInt(balanceUpdatedTimeStr, 10, 64); err == nil {
				channel.BalanceUpdatedTime = balanceUpdatedTime
			}
		}

		// 渠道信息 JSON
		if channelInfoJSON := getCell(28); channelInfoJSON != "" {
			var channelInfo model.ChannelInfo
			if err := json.Unmarshal([]byte(channelInfoJSON), &channelInfo); err == nil {
				channel.ChannelInfo = channelInfo
			}
		}

		channels = append(channels, channel)
	}

	return &ChannelExcelParseResult{
		Channels: channels,
		Errors:   errors,
	}, nil
}
