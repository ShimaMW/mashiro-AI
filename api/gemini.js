// ============================================================================
// Vercel Serverless Function: /api/gemini
// 環境変数 GEMINI_API_KEY が設定されている場合、APIキーをフロントに露出させずに
// Gemini Files API (Resumable Upload) および Gemini 生成（カイポケ用・iBow用）を実行します
// ============================================================================

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method Not Allowed' });
  }

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return res.status(400).json({
      error: 'Vercel環境変数 GEMINI_API_KEY が未設定です。画面右上の設定からAPIキーを設定するか、Vercelに環境変数を追加してください。',
      code: 'NO_ENV_KEY'
    });
  }

  try {
    const { action, payload = {} } = req.body || {};

    // 1. 接続確認
    if (action === 'ping') {
      return res.status(200).json({
        success: true,
        hasKey: true,
        message: 'Vercel Gemini API エンドポイントが有効です'
      });
    }

    // 2. Resumable Upload セッションURLの発行（大容量音声用）
    if (action === 'initUpload') {
      const { mimeType = 'audio/mp3', numBytes, displayName = 'meeting_audio', origin } = payload;
      const clientOrigin = req.headers.origin || origin || 'https://mashiro-ai.vercel.app';
      const startUploadRes = await fetch(
        `https://generativelanguage.googleapis.com/upload/v1beta/files?key=${apiKey}`,
        {
          method: 'POST',
          headers: {
            'Origin': clientOrigin,
            'X-Goog-Upload-Protocol': 'resumable',
            'X-Goog-Upload-Command': 'start',
            'X-Goog-Upload-Header-Content-Length': String(numBytes),
            'X-Goog-Upload-Header-Content-Type': mimeType,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({ file: { display_name: displayName } })
        }
      );

      if (!startUploadRes.ok) {
        const errText = await startUploadRes.text();
        throw new Error(`Upload初期化失敗: ${errText}`);
      }

      const uploadUrl = startUploadRes.headers.get('x-goog-upload-url');
      return res.status(200).json({ uploadUrl });
    }

    // 3. 議事録生成リクエスト
    if (action === 'generate') {
      const {
        prompt,
        fileUri,
        mimeType,
        modelName = 'gemini-3.6-flash',
        fileNameToCleanUp,
        responseSchema
      } = payload;

      const parts = [{ text: prompt }];
      if (fileUri) {
        parts.push({
          file_data: {
            mime_type: mimeType || 'audio/mp3',
            file_uri: fileUri
          }
        });
      }

      const defaultSchema = {
        type: 'OBJECT',
        properties: {
          summary: { type: 'STRING', description: '会議の主な論点、背景、本人・家族の意向、決定事項を300〜500文字程度でしっかり要約' },
          item1: { type: 'STRING', description: '開催目的と、検討した議題を①②③とナンバリングして記載し、振り返りがしやすいように各議題100文字以内で要約' },
          item2: { type: 'STRING', description: '議題に対応させ、どの職種（誰）がどのような発言・専門的見解を述べたかを詳細な文章で記載' },
          item3: { type: 'STRING', description: '話し合って合意・決定された内容や、決定に至った具体的な理由、ましろの役割やToDoを詳細に記載' },
          item4: { type: 'STRING', description: '解決できなかった内容、今後の観察・モニタリングポイント、次回の開催予定などを記載' }
        },
        required: ['summary', 'item1', 'item2', 'item3', 'item4']
      };

      const geminiRes = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            contents: [{ parts }],
            generationConfig: {
              responseMimeType: 'application/json',
              temperature: 0.2,
              responseSchema: responseSchema || defaultSchema
            }
          })
        }
      );

      const json = await geminiRes.json();

      // アップロード済み一時音声ファイルの即時削除（プライバシー保護）
      if (fileNameToCleanUp) {
        fetch(`https://generativelanguage.googleapis.com/v1beta/${fileNameToCleanUp}?key=${apiKey}`, {
          method: 'DELETE'
        }).catch(() => {});
      }

      if (json.error) {
        return res.status(500).json({ error: `【AI解析エラー】: ${json.error.message}` });
      }
      if (!json.candidates || !json.candidates[0]?.content?.parts?.[0]?.text) {
        return res.status(500).json({ error: '【AI応答エラー】: AIからの有効なテキスト応答がありませんでした。' });
      }

      const resultData = JSON.parse(json.candidates[0].content.parts[0].text);
      return res.status(200).json(resultData);
    }

    return res.status(400).json({ error: `Unknown action: ${action}` });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
