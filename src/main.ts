// 型定義
interface DonationDetails {
  date: string;
  name: string;
  amount: number;
  frequency: string;
}

interface AppConfig {
  slackWebhookUrl: string | null;
  spreadsheetId: string | null;
  sheetName: string | null;
}

interface SlackWorkflowPayload {
  date: string;
  amount: string;
  frequency: string;
}

// 日付ユーティリティ
class DateUtils {
  /**
   * 日付文字列を正規化する
   */
  public static normalize(dateString: string): string {
    const match = dateString.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})/);
    if (!match) return dateString;

    const [, year, month, day, hour, minute, second] = match;
    return `${year}/${month.padStart(2, '0')}/${day.padStart(2, '0')} ${hour.padStart(2, '0')}:${minute.padStart(2, '0')}:${second.padStart(2, '0')}`;
  }

  /**
   * 日付文字列をDateオブジェクトに変換する
   */
  public static parseDate(dateString: string): Date | null {
    const date = new Date(dateString);
    return isNaN(date.getTime()) ? null : date;
  }
}

// 設定サービス
class ConfigService {
  private readonly config: AppConfig;

  constructor() {
    this.config = this.loadConfig();
  }

  private loadConfig(): AppConfig {
    const properties = PropertiesService.getScriptProperties();
    return {
      slackWebhookUrl: properties.getProperty('SLACK_WEBHOOK_URL'),
      spreadsheetId: properties.getProperty('GOOGLE_SHEETS_ID'),
      sheetName: properties.getProperty('GOOGLE_SHEETS_NAME'),
    };
  }

  public getConfig(): AppConfig {
    return this.config;
  }

  public isConfigValid(): boolean {
    return Boolean(this.config.slackWebhookUrl);
  }

  public isSheetsConfigValid(): boolean {
    return Boolean(this.config.spreadsheetId && this.config.sheetName);
  }
}

// Gmail サービス
class GmailService {
  private static readonly SEARCH_QUERY =
    'subject:"【Syncable】新規の支援を受け付けました。" is:unread';

  public static processNewDonations(
    onDonationFound: (details: DonationDetails, messageId: string) => void
  ): void {
    const threads = GmailApp.search(this.SEARCH_QUERY);
    const messages = GmailApp.getMessagesForThreads(threads).flat();

    for (const message of messages) {
      try {
        const details = this.extractDonationDetails(message.getPlainBody());
        if (details) {
          onDonationFound(details, message.getId());
        } else {
          console.warn(`寄付情報の抽出に失敗しました (messageId: ${message.getId()})`);
        }
      } catch (error) {
        console.error(`メッセージ処理中にエラーが発生しました: ${error}`);
      }
    }
  }

  private static extractDonationDetails(body: string): DonationDetails | null {
    const patterns = {
      date: /支援受付日時:\s*(.+)/,
      name: /支援付者名:\s*(.+)/,
      amount: /支援金額:\s*([\d,]+)\s*円/,
      frequency: /支援頻度:\s*(.+)/,
    };

    const matches = {
      date: body.match(patterns.date),
      name: body.match(patterns.name),
      amount: body.match(patterns.amount),
      frequency: body.match(patterns.frequency),
    };

    if (!Object.values(matches).every((match) => match)) {
      return null;
    }

    // 日付を正規化
    const rawDate = matches.date![1].trim();
    const normalizedDate = DateUtils.normalize(rawDate);

    // 金額をnumberに変換
    const amountString = matches.amount![1].trim();
    const amount = parseInt(amountString.replace(/,/g, ''), 10);
    if (isNaN(amount)) {
      return null;
    }

    return {
      date: normalizedDate,
      name: matches.name![1].split(/\s{2,}/)[0].trim(),
      amount: amount,
      frequency: matches.frequency![1].split(' ')[0].trim(),
    };
  }
}

// Slack サービス
class SlackService {
  private webhookUrl: string;

  constructor(webhookUrl: string) {
    this.webhookUrl = webhookUrl;
  }

  public sendDonationNotification(date: string, amount: number, frequency: string): void {
    const payload: SlackWorkflowPayload = {
      date: date,
      amount: `${amount.toLocaleString()}円`,
      frequency: frequency,
    };

    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
    };

    try {
      const response = UrlFetchApp.fetch(this.webhookUrl, options);
      console.log(`Slack通知を送信しました: ${date}, ${amount}円, ${frequency}`);
    } catch (error) {
      console.error(`Slackへの通知に失敗しました: ${error}`);
      throw error;
    }
  }
}

// Google Sheets サービス
class SheetsService {
  private spreadsheetId: string;
  private sheetName: string;

  constructor(spreadsheetId: string, sheetName: string) {
    this.spreadsheetId = spreadsheetId;
    this.sheetName = sheetName;
  }

  public recordDonation(date: string, name: string, amount: number, frequency: string): void {
    try {
      const sheet = this.getSheet();
      if (!sheet) {
        throw new Error(`シートが見つかりません: ${this.sheetName}`);
      }

      // 日付文字列をDateに変換
      const parsedDate = DateUtils.parseDate(date);
      const dateValue = parsedDate || date;

      sheet.appendRow([dateValue, name, amount, frequency]);
      console.log(`Google Sheetsに記録しました: ${name}, ${amount}円`);
    } catch (error) {
      console.error(`Google Sheetsへの記録に失敗しました: ${error}`);
      throw error;
    }
  }

  private getSheet(): GoogleAppsScript.Spreadsheet.Sheet | null {
    try {
      const spreadsheet = SpreadsheetApp.openById(this.spreadsheetId);
      return spreadsheet.getSheetByName(this.sheetName);
    } catch (error) {
      console.error(`Google Sheetsの取得に失敗しました: ${error}`);
      return null;
    }
  }

  public ensureHeaders(): void {
    try {
      const sheet = this.getSheet();
      if (!sheet) return;

      const lastRow = sheet.getLastRow();
      if (lastRow === 0) {
        sheet.appendRow(['日時', '寄付者名', '金額', '頻度']);
        console.log('ヘッダーを追加しました');
      }
    } catch (error) {
      console.error(`ヘッダーの確認/追加に失敗しました: ${error}`);
    }
  }
}

// メインアプリケーション
class DonationNotifierApp {
  private configService: ConfigService;
  private slackService: SlackService | null = null;
  private sheetsService: SheetsService | null = null;

  constructor() {
    this.configService = new ConfigService();
    this.initializeServices();
  }

  private initializeServices(): void {
    const config = this.configService.getConfig();

    if (config.slackWebhookUrl) {
      this.slackService = new SlackService(config.slackWebhookUrl);
    }

    if (config.spreadsheetId && config.sheetName) {
      this.sheetsService = new SheetsService(config.spreadsheetId, config.sheetName);
      this.sheetsService.ensureHeaders();
    }
  }

  public checkForNewDonations(): void {
    if (!this.configService.isConfigValid()) {
      console.warn('Slack Webhook URL が設定されていません');
      return;
    }

    console.log('新規寄付通知のチェックを開始します');

    try {
      GmailService.processNewDonations((details: DonationDetails, messageId: string) => {
        this.processDonation(details, messageId);
      });

      console.log('寄付通知のチェックが完了しました');
    } catch (error) {
      console.error('寄付通知のチェック中にエラーが発生しました:', error);
    }
  }

  private processDonation(details: DonationDetails, messageId: string): void {
    const { date, name, amount, frequency } = details;

    console.log(`寄付情報を処理中: ${date}, ${name}, ${amount}円, ${frequency}`);

    try {
      if (this.slackService) {
        this.slackService.sendDonationNotification(date, amount, frequency);
      }

      if (this.sheetsService) {
        this.sheetsService.recordDonation(date, name, amount, frequency);
      }

      // すべての処理が成功した場合にのみ、メールを既読にする
      GmailApp.getMessageById(messageId).markRead();
      console.log(`寄付情報の処理が完了しました (messageId: ${messageId})`);
    } catch (error) {
      console.error(`寄付情報の処理中にエラーが発生しました (messageId: ${messageId}):`, error);
    }
  }
}

// GAS用のグローバル関数
function checkForNewDonations(): void {
  const app = new DonationNotifierApp();
  app.checkForNewDonations();
}
