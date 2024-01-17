require 'write_xlsx'
require 'logger'

# ロガーの初期化
logger = Logger.new("data_collection.log")

# ユーザー入力データの取得
def get_user_input(input_history, logger)
  print "データを入力してください（あるいは「exit」と入力して終了し、「save」と入力して保存する。）: "
  input = gets.chomp.strip
  input_history << input unless input.downcase == 'exit' || input.downcase == 'save'
  logger.info("ユーザーの入力。: #{input}")
  input
end

# 入力が数値かどうかをチェックする
def valid_input?(input)
  input.match?(/^\d+(\.\d+)?$/) || %w[exit save].include?(input.downcase)
end

# グラフを含むデータをexcelに保存
def save_data(data, logger)
  return if data.empty?

  file_name = "measurement_data_#{Time.now.strftime('%Y%m%d%H%M%S')}.xlsx"
  workbook = WriteXLSX.new(file_name)
  worksheet = workbook.add_worksheet
  chart = workbook.add_chart(type: 'line', embedded: 1)

 worksheet.write_row('A1', ['time (s)', 'data'])
  data.each_with_index do |row, i|
    worksheet.write_row(i+1, 0, row)
  end

  # チャートのデータ範囲を設定する
  chart.add_series(
    categories: "=Sheet1!$A$2:$A$#{data.length + 1}",
    values:     "=Sheet1!$B$2:$B$#{data.length + 1}"
  )

  # チャートのタイトルと軸ラベルの設定
  chart.set_title(name: '測定データグラフ')
  chart.set_x_axis(name: '時間 (s)')
  chart.set_y_axis(name: 'データ')

  # チャート挿入
  worksheet.insert_chart('D2', chart)

  workbook.close
  logger.info("データをファイルに保存する: #{file_name}")
  puts "データがに保存された #{file_name}"
end

# メインプログラム
def main
  data = []
  input_history = []
  logger = Logger.new('measurement.log')
  logger.info("プログラム起動")

  puts "データ収集プロセスへようこそ！"
  puts "'exit' と入力すればいつでもプログラムを終了でき、'save' と入力すればデータを保存できる。	"

  start_time = Time.now

  loop do
    input = get_user_input(input_history, logger)
    break if input.downcase == 'exit'

    if input.downcase == 'save'
      save_data(data, logger)
      data.clear
      next
    end

    unless valid_input?(input)
      puts "入力が無効です。"
      logger.error("無効な入力")
      next
    end

    elapsed_time = Time.now - start_time
    data << [elapsed_time.round(1), input.to_f]
  end

  unless data.empty?
    print "未保存のデータがあります。 保存されていますか？ (yes／no）。 "
    save_data(data, logger) if gets.chomp.strip.downcase == 'yes'
  end

  logger.info("プログラム終了")
  puts "ありがとうございます！"
end

main
