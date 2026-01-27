// 初期化
document.addEventListener('DOMContentLoaded', () => {
  loadCategories();
  loadProducts();
  loadStats();
});

// カテゴリ読み込み
async function loadCategories() {
  const res = await fetch('/api/categories');
  const categories = await res.json();

  const filterSelect = document.getElementById('categoryFilter');
  const formSelect = document.getElementById('category');

  categories.forEach(cat => {
    filterSelect.innerHTML += `<option value="${cat.name}">${cat.name}</option>`;
    formSelect.innerHTML += `<option value="${cat.name}">${cat.name}</option>`;
  });
}

// 統計情報読み込み
async function loadStats() {
  const res = await fetch('/api/stats');
  const stats = await res.json();

  document.getElementById('totalProducts').textContent = stats.totalProducts;
  document.getElementById('totalStock').textContent = stats.totalStock;
  document.getElementById('totalValue').textContent = `¥${stats.totalValue.toLocaleString()}`;
  document.getElementById('lowStock').textContent = stats.lowStock;
}

// 商品一覧読み込み
async function loadProducts() {
  const search = document.getElementById('searchInput').value;
  const category = document.getElementById('categoryFilter').value;

  const params = new URLSearchParams();
  if (search) params.append('search', search);
  if (category) params.append('category', category);

  const res = await fetch(`/api/products?${params}`);
  const products = await res.json();

  const tbody = document.getElementById('productsBody');
  tbody.innerHTML = '';

  if (products.length === 0) {
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;color:#666;">商品がありません</td></tr>';
    return;
  }

  products.forEach(p => {
    const stockClass = p.stock_quantity <= 3 ? 'stock-low' : '';
    tbody.innerHTML += `
      <tr>
        <td>${escapeHtml(p.name)}</td>
        <td>${escapeHtml(p.brand || '-')}</td>
        <td>${escapeHtml(p.category)}</td>
        <td>${escapeHtml(p.size || '-')}</td>
        <td>${escapeHtml(p.condition)}</td>
        <td>¥${p.purchase_price.toLocaleString()}</td>
        <td>¥${p.selling_price.toLocaleString()}</td>
        <td class="${stockClass}">${p.stock_quantity}</td>
        <td class="actions">
          <button class="btn btn-success btn-sm" onclick="showStockModal(${p.id}, 'in', '${escapeHtml(p.name)}', ${p.stock_quantity})">入庫</button>
          <button class="btn btn-sm" onclick="showStockModal(${p.id}, 'out', '${escapeHtml(p.name)}', ${p.stock_quantity})">出庫</button>
          <button class="btn btn-sm" onclick="editProduct(${p.id})">編集</button>
          <button class="btn btn-danger btn-sm" onclick="deleteProduct(${p.id})">削除</button>
        </td>
      </tr>
    `;
  });
}

// HTMLエスケープ
function escapeHtml(str) {
  if (!str) return '';
  return str.replace(/[&<>"']/g, char => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  }[char]));
}

// 商品登録モーダル表示
function showAddModal() {
  document.getElementById('modalTitle').textContent = '商品登録';
  document.getElementById('productForm').reset();
  document.getElementById('productId').value = '';
  document.getElementById('stockQuantityGroup').style.display = 'block';
  document.getElementById('productModal').style.display = 'block';
}

// 商品編集
async function editProduct(id) {
  const res = await fetch(`/api/products/${id}`);
  const p = await res.json();

  document.getElementById('modalTitle').textContent = '商品編集';
  document.getElementById('productId').value = p.id;
  document.getElementById('name').value = p.name;
  document.getElementById('brand').value = p.brand || '';
  document.getElementById('category').value = p.category;
  document.getElementById('size').value = p.size || '';
  document.getElementById('condition').value = p.condition;
  document.getElementById('purchase_price').value = p.purchase_price;
  document.getElementById('selling_price').value = p.selling_price;
  document.getElementById('notes').value = p.notes || '';
  document.getElementById('stockQuantityGroup').style.display = 'none';
  document.getElementById('productModal').style.display = 'block';
}

// 商品保存
async function saveProduct(e) {
  e.preventDefault();

  const id = document.getElementById('productId').value;
  const data = {
    name: document.getElementById('name').value,
    brand: document.getElementById('brand').value,
    category: document.getElementById('category').value,
    size: document.getElementById('size').value,
    condition: document.getElementById('condition').value,
    purchase_price: parseInt(document.getElementById('purchase_price').value),
    selling_price: parseInt(document.getElementById('selling_price').value),
    notes: document.getElementById('notes').value
  };

  if (!id) {
    data.stock_quantity = parseInt(document.getElementById('stock_quantity').value) || 0;
  }

  const url = id ? `/api/products/${id}` : '/api/products';
  const method = id ? 'PUT' : 'POST';

  await fetch(url, {
    method,
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data)
  });

  closeModal();
  loadProducts();
  loadStats();
}

// 商品削除
async function deleteProduct(id) {
  if (!confirm('この商品を削除しますか？')) return;

  await fetch(`/api/products/${id}`, { method: 'DELETE' });
  loadProducts();
  loadStats();
}

// モーダル閉じる
function closeModal() {
  document.getElementById('productModal').style.display = 'none';
}

// 入出庫モーダル表示
function showStockModal(id, type, name, currentStock) {
  document.getElementById('stockModalTitle').textContent = type === 'in' ? '入庫処理' : '出庫処理';
  document.getElementById('stockProductId').value = id;
  document.getElementById('stockType').value = type;
  document.getElementById('stockProductName').textContent = name;
  document.getElementById('currentStock').textContent = `現在の在庫: ${currentStock}個`;
  document.getElementById('stockQuantity').value = 1;
  document.getElementById('stockNotes').value = '';
  document.getElementById('stockModal').style.display = 'block';
}

// 入出庫処理
async function processStock(e) {
  e.preventDefault();

  const id = document.getElementById('stockProductId').value;
  const type = document.getElementById('stockType').value;
  const quantity = parseInt(document.getElementById('stockQuantity').value);
  const notes = document.getElementById('stockNotes').value;

  const endpoint = type === 'in' ? 'stock-in' : 'stock-out';

  const res = await fetch(`/api/products/${id}/${endpoint}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ quantity, notes })
  });

  if (!res.ok) {
    const error = await res.json();
    alert(error.error);
    return;
  }

  closeStockModal();
  loadProducts();
  loadStats();
}

// 入出庫モーダル閉じる
function closeStockModal() {
  document.getElementById('stockModal').style.display = 'none';
}

// Enterキーで検索
document.getElementById('searchInput').addEventListener('keypress', (e) => {
  if (e.key === 'Enter') loadProducts();
});
