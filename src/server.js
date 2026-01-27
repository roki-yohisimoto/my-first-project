const express = require('express');
const path = require('path');
const db = require('./database');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.static(path.join(__dirname, '..', 'public')));

// === カテゴリAPI ===
app.get('/api/categories', (req, res) => {
  const categories = db.prepare('SELECT * FROM categories ORDER BY name').all();
  res.json(categories);
});

// === 商品API ===
// 商品一覧取得（検索・フィルター対応）
app.get('/api/products', (req, res) => {
  const { search, category, brand } = req.query;
  let query = 'SELECT * FROM products WHERE 1=1';
  const params = [];

  if (search) {
    query += ' AND (name LIKE ? OR brand LIKE ? OR notes LIKE ?)';
    const searchTerm = `%${search}%`;
    params.push(searchTerm, searchTerm, searchTerm);
  }
  if (category) {
    query += ' AND category = ?';
    params.push(category);
  }
  if (brand) {
    query += ' AND brand LIKE ?';
    params.push(`%${brand}%`);
  }

  query += ' ORDER BY updated_at DESC';
  const products = db.prepare(query).all(...params);
  res.json(products);
});

// 商品詳細取得
app.get('/api/products/:id', (req, res) => {
  const product = db.prepare('SELECT * FROM products WHERE id = ?').get(req.params.id);
  if (!product) {
    return res.status(404).json({ error: '商品が見つかりません' });
  }
  res.json(product);
});

// 商品登録
app.post('/api/products', (req, res) => {
  const { name, brand, category, size, condition, purchase_price, selling_price, stock_quantity, notes } = req.body;

  if (!name || !category || !condition || purchase_price === undefined || selling_price === undefined) {
    return res.status(400).json({ error: '必須項目が不足しています' });
  }

  const stmt = db.prepare(`
    INSERT INTO products (name, brand, category, size, condition, purchase_price, selling_price, stock_quantity, notes)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);

  const result = stmt.run(name, brand || null, category, size || null, condition, purchase_price, selling_price, stock_quantity || 0, notes || null);

  const newProduct = db.prepare('SELECT * FROM products WHERE id = ?').get(result.lastInsertRowid);
  res.status(201).json(newProduct);
});

// 商品更新
app.put('/api/products/:id', (req, res) => {
  const { name, brand, category, size, condition, purchase_price, selling_price, notes } = req.body;

  const stmt = db.prepare(`
    UPDATE products
    SET name = ?, brand = ?, category = ?, size = ?, condition = ?,
        purchase_price = ?, selling_price = ?, notes = ?, updated_at = CURRENT_TIMESTAMP
    WHERE id = ?
  `);

  stmt.run(name, brand, category, size, condition, purchase_price, selling_price, notes, req.params.id);

  const updated = db.prepare('SELECT * FROM products WHERE id = ?').get(req.params.id);
  res.json(updated);
});

// 商品削除
app.delete('/api/products/:id', (req, res) => {
  db.prepare('DELETE FROM inventory_transactions WHERE product_id = ?').run(req.params.id);
  db.prepare('DELETE FROM products WHERE id = ?').run(req.params.id);
  res.json({ message: '削除しました' });
});

// === 入出庫API ===
// 入庫処理
app.post('/api/products/:id/stock-in', (req, res) => {
  const { quantity, notes } = req.body;
  const productId = req.params.id;

  if (!quantity || quantity <= 0) {
    return res.status(400).json({ error: '数量は1以上を指定してください' });
  }

  const transaction = db.transaction(() => {
    // 在庫数を増やす
    db.prepare('UPDATE products SET stock_quantity = stock_quantity + ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?')
      .run(quantity, productId);

    // 入庫履歴を記録
    db.prepare('INSERT INTO inventory_transactions (product_id, type, quantity, notes) VALUES (?, ?, ?, ?)')
      .run(productId, 'in', quantity, notes || null);
  });

  transaction();

  const product = db.prepare('SELECT * FROM products WHERE id = ?').get(productId);
  res.json(product);
});

// 出庫処理
app.post('/api/products/:id/stock-out', (req, res) => {
  const { quantity, notes } = req.body;
  const productId = req.params.id;

  if (!quantity || quantity <= 0) {
    return res.status(400).json({ error: '数量は1以上を指定してください' });
  }

  const product = db.prepare('SELECT stock_quantity FROM products WHERE id = ?').get(productId);
  if (product.stock_quantity < quantity) {
    return res.status(400).json({ error: '在庫が不足しています' });
  }

  const transaction = db.transaction(() => {
    // 在庫数を減らす
    db.prepare('UPDATE products SET stock_quantity = stock_quantity - ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?')
      .run(quantity, productId);

    // 出庫履歴を記録
    db.prepare('INSERT INTO inventory_transactions (product_id, type, quantity, notes) VALUES (?, ?, ?, ?)')
      .run(productId, 'out', quantity, notes || null);
  });

  transaction();

  const updatedProduct = db.prepare('SELECT * FROM products WHERE id = ?').get(productId);
  res.json(updatedProduct);
});

// 入出庫履歴取得
app.get('/api/products/:id/transactions', (req, res) => {
  const transactions = db.prepare(`
    SELECT * FROM inventory_transactions
    WHERE product_id = ?
    ORDER BY created_at DESC
  `).all(req.params.id);
  res.json(transactions);
});

// === 統計API ===
app.get('/api/stats', (req, res) => {
  const totalProducts = db.prepare('SELECT COUNT(*) as count FROM products').get().count;
  const totalStock = db.prepare('SELECT SUM(stock_quantity) as total FROM products').get().total || 0;
  const totalValue = db.prepare('SELECT SUM(stock_quantity * selling_price) as total FROM products').get().total || 0;
  const lowStock = db.prepare('SELECT COUNT(*) as count FROM products WHERE stock_quantity <= 3').get().count;

  res.json({
    totalProducts,
    totalStock,
    totalValue,
    lowStock
  });
});

app.listen(PORT, () => {
  console.log(`在庫管理サーバーが起動しました: http://localhost:${PORT}`);
});
