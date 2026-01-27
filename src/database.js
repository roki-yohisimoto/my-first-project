const Database = require('better-sqlite3');
const path = require('path');

const db = new Database(path.join(__dirname, '..', 'inventory.db'));

// テーブル作成
db.exec(`
  -- 商品テーブル
  CREATE TABLE IF NOT EXISTS products (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    brand TEXT,
    category TEXT NOT NULL,
    size TEXT,
    condition TEXT NOT NULL,
    purchase_price INTEGER NOT NULL,
    selling_price INTEGER NOT NULL,
    stock_quantity INTEGER DEFAULT 0,
    notes TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
  );

  -- 入出庫履歴テーブル
  CREATE TABLE IF NOT EXISTS inventory_transactions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    product_id INTEGER NOT NULL,
    type TEXT NOT NULL CHECK(type IN ('in', 'out')),
    quantity INTEGER NOT NULL,
    notes TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (product_id) REFERENCES products(id)
  );

  -- カテゴリマスタ
  CREATE TABLE IF NOT EXISTS categories (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE
  );

  -- 初期カテゴリデータ
  INSERT OR IGNORE INTO categories (name) VALUES
    ('トップス'),
    ('ボトムス'),
    ('アウター'),
    ('ワンピース'),
    ('バッグ'),
    ('シューズ'),
    ('アクセサリー'),
    ('その他');
`);

module.exports = db;
