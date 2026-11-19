export const translations = {
	'pptx.options.addIns.mtxDecompressor': 'MicroType Express フォント展開ツール',
	'pptx.options.addIns.mtxDecompressorDescription':
		'埋め込まれた MicroType Express フォントを展開します。',
	'pptx.options.addIns.collaboration': 'リアルタイム共同作業',
	'pptx.options.addIns.collaborationDescription':
		'CRDT ベースの共同編集、プレゼンス、同期リレーをサポートします。',
	'pptx.options.addIns.locales': '言語パック',
	'pptx.options.addIns.localesDescription': 'ビューアーの画面に使用できる追加の表示言語です。',
	'pptx.options.trust.label': 'トラスト センター',
	'pptx.options.trust.description': 'ドキュメントの安全性とセキュリティを保護します。',
	'pptx.options.trust.settingsTitle': 'セキュリティ設定',
	'pptx.options.trust.settingsDescription':
		'これらの設定は、開いたコンテンツの安全性を保つのに役立ちます。既定の設定をお勧めします。',
	'pptx.options.trust.protectedView': 'プレゼンテーションを保護ビューで開く',
	'pptx.options.trust.protectedViewInfo':
		'ドキュメントは読み取り専用で開かれます。変更するには [編集を有効にする] を使用します。',
	'pptx.options.trust.allowExternalContent': '外部コンテンツ (リモートの画像とメディア) を許可する',
	'pptx.options.trust.allowExternalContentInfo':
		'オフ (既定、PowerPoint と同じ) の場合、リンクされた画像やメディアはダウンロードされず、プレゼンテーションごとにセキュリティ バーで [コンテンツの有効化] を選べます。',
	'pptx.options.trust.allowFontDownloads':
		'インストールも埋め込みもされていないフォントをダウンロードする (Google Fonts)',
	'pptx.options.trust.allowFontDownloadsInfo':
		'オンの場合、不足しているフォントを取得するために、ブラウザーが fonts.googleapis.com と fonts.gstatic.com (または組織が設定したフォント サーバー) に接続します。オフの場合、テキストにはインストール済み、埋め込み、または代替のフォントが使用されるため、改行位置が PowerPoint と異なる場合があります。',
	'pptx.options.trust.confirmHyperlinks': '外部ハイパーリンクを開く前に確認する',
} as const;
