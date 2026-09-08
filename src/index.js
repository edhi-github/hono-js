import { Hono } from 'hono';
import { cors } from 'hono/cors';
import { serveStatic } from 'hono/cloudflare-workers';
import midtransClient from 'midtrans-client';
import crypto from 'node:crypto';
import { S3Client, PutObjectCommand } from '@aws-sdk/client-s3';
import ExcelJS from 'exceljs';
import { GoogleGenerativeAI } from '@google/generative-ai';

const app = new Hono();

// Middleware CORS
app.use('*', cors({
    origin: '*',
    allowMethods: ['GET', 'POST', 'PUT', 'PATCH', 'DELETE', 'OPTIONS'],
    allowHeaders: ['Content-Type', 'Authorization', 'x-shop-slug', 'x-shop-id', 'Accept']
}));

// Helper untuk koneksi Cloudflare D1 Database
function getDbPool(c) {
   return c.env.beda_pos; 
}

// Helper S3 Client untuk R2
function getS3Client(c) {
    const env = c.env;
    return new S3Client({
        region: 'auto',
        endpoint: env.R2_ENDPOINT || '',
        credentials: {
            accessKeyId: env.R2_ACCESS_KEY_ID || '',
            secretAccessKey: env.R2_SECRET_ACCESS_KEY || ''
        }
    });
}

// Helper Snap Midtrans
function getSnapClient(c) {
    const env = c.env;
    return new midtransClient.Snap({
        isProduction: env.MIDTRANS_IS_PRODUCTION === 'true',
        serverKey: env.MIDTRANS_SERVER_KEY || '',
        clientKey: env.MIDTRANS_CLIENT_KEY || ''
    });
}

async function getShopIdBySlug(db, slug) {
    if (!slug) return null;
    const { results } = await db.prepare('SELECT id FROM shops WHERE slug = ?').bind(slug).all();
    return results && results.length > 0 ? results[0].id : null;
}

// Middleware Verifikasi Akses Warung
async function verifikasiAksesWarung(c, next) {
    const shopSlug = c.req.query('shop') || c.req.header('x-shop-slug');
    if (!shopSlug) {
        let bodyShop = null;
        try {
            const body = await c.req.parseBody();
            bodyShop = body.shop;
        } catch (e) {}
        if (!bodyShop) {
            return c.json({ success: false, message: "Parameter warung/shop tidak ditemukan." }, 400);
        }
    }
    await next();
}

// Middleware Cek Masa Aktif Sub
async function cekMasaAktifSub(c, next) {
    try {
        const pool = getDbPool(c);
        let shopId = c.req.header('x-shop-id') || c.req.query('shop_id');
        let bodyShop = null;

        try {
            const body = await c.req.parseBody();
            bodyShop = body.shop;
        } catch (e) {}

        const shopSlug = c.req.query('shop') || c.req.header('x-shop-slug') || bodyShop;

        if ((!shopId || shopId === 'null' || shopId === 'undefined') && shopSlug) {
            shopId = await getShopIdBySlug(pool, shopSlug);
        }

        if (!shopId) {
            return c.json({ success: false, message: "Shop ID atau Parameter Shop tidak ditemukan/valid" }, 400);
        }

        const { results: shops } = await pool.prepare(
            `SELECT s.subscription_until, s.max_transactions_monthly, s.package_id, p.slug AS package_slug 
             FROM shops s 
             LEFT JOIN packages p ON s.package_id = p.id 
             WHERE s.id = ?`
        ).bind(shopId).all();

        if (!shops || shops.length === 0) {
            return c.json({ success: false, message: "Toko tidak ditemukan" }, 404);
        }

        const shop = shops[0];
        const sekarang = new Date();
        const subUntil = new Date(shop.subscription_until);

        const batasToleransi = new Date(subUntil);
        batasToleransi.setDate(batasToleransi.getDate() + 1);

        const isExpiredByDate = sekarang > batasToleransi;
        const isSultan = (shop.package_id === 3 || (shop.package_slug && shop.package_slug.toLowerCase() === 'sultan'));
        const remainingQuota = shop.max_transactions_monthly || 0;
        const isQuotaExhausted = !isSultan && (remainingQuota <= 0);

        if (isExpiredByDate || isQuotaExhausted) {
            if (isExpiredByDate && shop.max_transactions_monthly > 0) {
                await pool.prepare("UPDATE shops SET max_transactions_monthly = 0, subscription_status = 'expired' WHERE id = ?").bind(shopId).run();
            }

            const pesanError = isExpiredByDate 
                ? "Masa aktif langganan Anda telah habis. Silakan lakukan perpanjangan paket!" 
                : "Kuota transaksi bulanan Anda telah habis. Silakan perpanjang atau tingkatkan paket Anda!";

            return c.json({ 
                success: false, 
                is_expired: true,
                message: pesanError 
            }, 403);
        }

        c.set('shop', shop);
        await next();
    } catch (error) {
        console.error("Error pada cekMasaAktifSub:", error);
        return c.json({ success: false, message: "Error validasi langganan: " + error.message }, 500);
    }
}

function getTimeZoneByWilayah(wilayah) {
    if (!wilayah) return 'Asia/Jakarta';
    const w = wilayah.toLowerCase();
    
    const witaProvinces = [
        'bali', 'nusa tenggara barat', 'nusa tenggara timur', 
        'kalimantan selatan', 'kalimantan timur', 'kalimantan utara',
        'sulawesi utara', 'sulawesi tengah', 'sulawesi selatan', 
        'sulawesi tenggara', 'gorontalo', 'sulawesi barat'
    ];
    if (witaProvinces.some(prov => w.includes(prov))) return 'Asia/Makassar';

    const witProvinces = [
        'maluku', 'maluku utara', 'papua', 'papua barat', 
        'papua selatan', 'papua tengah', 'papua pegunungan', 'papua barat daya'
    ];
    if (witProvinces.some(prov => w.includes(prov))) return 'Asia/Jayapura';

    return 'Asia/Jakarta';
}

// ---------------- INFO SHOPS ----------------
app.get('/api/shops/info', verifikasiAksesWarung, async (c) => {
    try {
        const pool = getDbPool(c);
        const shopSlug = c.req.query('shop') || c.req.header('x-shop-slug');
        const { results: rows } = await pool.prepare(
            `SELECT id, shop_name, has_tax, tax_percentage, discount_percentage, show_cash_payment, bank_rekening_info, qris_image_url 
             FROM shops WHERE slug = ?`
        ).bind(shopSlug).all();

        if (!rows || rows.length === 0) {
            return c.json({ success: false, message: "Warung tidak ditemukan." }, 404);
        }

        return c.json({ success: true, data: rows[0] });
    } catch (error) {
        console.error("Error ambil data toko:", error);
        return c.json({ success: false, message: "Gagal mengambil data toko." }, 500);
    }
});

// ---------------- PRODUK & STOK ----------------
app.post('/api/products/add-stock', verifikasiAksesWarung, cekMasaAktifSub, async (c) => {
    const pool = getDbPool(c);
    try {
        const body = await c.req.json();
        const shopSlug = body.shop || c.req.query('shop') || c.req.header('x-shop-slug');
        
        const productId = parseInt(body.product_id || body.id, 10);
        const qty = parseInt(body.qty, 10);
        const buyPrice = parseFloat(body.buy_price) || 0;

        const subtotalPO = parseFloat(body.subtotal_po) || (qty * buyPrice);
        const taxPO = parseFloat(body.tax_po) || 0;
        const otherCostPO = parseFloat(body.other_cost_po) || 0;
        const notes = body.notes || 'Penambahan Stok Manual';

        if (!productId || isNaN(productId) || isNaN(qty) || qty <= 0) {
            return c.json({ 
                success: false, 
                message: 'Data tidak valid. Produk ID dan jumlah stok (qty) wajib diisi.' 
            }, 400);
        }

        const shopId = await getShopIdBySlug(pool, shopSlug);
        if (!shopId) {
            return c.json({ success: false, message: 'Warung tidak ditemukan.' }, 404);
        }

        const totalBeliItem = qty * buyPrice;
        const totalBiayaTambahanPO = taxPO + otherCostPO;
        const bobotItem = subtotalPO > 0 ? (totalBeliItem / subtotalPO) : 1;
        const bebanTambahanPerPcs = qty > 0 ? ((bobotItem * totalBiayaTambahanPO) / qty) : 0;
        const effectiveBuyPrice = buyPrice + bebanTambahanPerPcs;

        const { results: prodRows } = await pool.prepare(
            'SELECT id, stock, cost_price, name FROM products WHERE id = ? AND shop_id = ?'
        ).bind(productId, shopId).all();

        if (!prodRows || prodRows.length === 0) {
            return c.json({ success: false, message: 'Produk tidak ditemukan.' }, 404);
        }

        const currentStock = prodRows[0].stock;
        const currentCostPrice = parseFloat(prodRows[0].cost_price) || 0;
        const newStock = currentStock + qty;

        let newCostPrice = currentCostPrice;
        if (newStock > 0) {
            const totalValLama = currentStock * currentCostPrice;
            const totalValBaru = qty * effectiveBuyPrice;
            newCostPrice = Math.round((totalValLama + totalValBaru) / newStock);
        } else {
            newCostPrice = Math.round(effectiveBuyPrice);
        }

        await pool.prepare('UPDATE products SET stock = ?, cost_price = ? WHERE id = ?').bind(newStock, newCostPrice, productId).run();

        await pool.prepare(
            `INSERT INTO stock_mutations (shop_id, product_id, type, qty, buy_price, subtotal_po, tax_po, other_cost_po, unit_cost, stock_before, stock_after, reference_number, notes) 
             VALUES (?, ?, 'IN', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(shopId, productId, qty, effectiveBuyPrice, subtotalPO, taxPO, otherCostPO, newCostPrice, currentStock, newStock, `RESTOCK-${Date.now()}`, notes).run();

        return c.json({
            success: true,
            message: `Stok bertambah +${qty}. HPP Efektif: Rp ${Math.round(effectiveBuyPrice).toLocaleString('id-ID')} | HPP Rata-rata baru: Rp ${newCostPrice.toLocaleString('id-ID')}`,
            new_stock: newStock,
            new_cost_price: newCostPrice
        });
    } catch (error) {
        console.error('Error tambah stok:', error);
        return c.json({ success: false, message: 'Gagal menambah stok: ' + error.message }, 500);
    }
});

app.get('/api/products', verifikasiAksesWarung, async (c) => {
    try {
        const pool = getDbPool(c);
        const shopSlug = c.req.query('shop') || c.req.header('x-shop-slug');
        const shopId = await getShopIdBySlug(pool, shopSlug);
        if (!shopId) {
            return c.json({ success: false, message: "Warung tidak ditemukan." }, 404);
        }

        const { results: rows } = await pool.prepare(
            `SELECT id, name, barcode, price, cost_price, discount_percentage, category, image_url, description, is_available, stock 
            FROM products 
            WHERE shop_id = ? AND is_active = 1 
            ORDER BY category ASC, name ASC`
        ).bind(shopId).all();

        return c.json({ success: true, data: rows });
    } catch (error) {
        console.error("Error mengambil data produk:", error);
        return c.json({ success: false, message: "Gagal mengambil daftar produk." }, 500);
    }
});

app.post('/api/products', verifikasiAksesWarung, cekMasaAktifSub, async (c) => {
    try {
        const pool = getDbPool(c);
        const s3 = getS3Client(c);
        const body = await c.req.parseBody();

        const shopSlug = c.req.query('shop') || c.req.header('x-shop-slug') || body.shop;
        const nama_produk = body.nama_produk || body.name;
        const harga = body.harga || body.price;
        const cost_price = parseFloat(body.cost_price) || 0;
        const kategori = body.kategori || body.category;
        const deskripsi = body.deskripsi || body.description || '';
        const stock = body.stock;
        const barcode = body.barcode || null;
        
        const shopId = await getShopIdBySlug(pool, shopSlug);
        if (!shopId) {
            return c.json({ success: false, message: 'Warung tidak terdaftar atau parameter shop tidak valid.' }, 404);
        }

        let urlFoto = '';
        const file = body.foto_produk;
        if (file && typeof file === 'object' && file.name) {
            const fileExtension = file.name.split('.').pop();
            const uniqueFilename = `product-${Date.now()}-${crypto.randomBytes(4).toString('hex')}.${fileExtension}`;
            const arrayBuffer = await file.arrayBuffer();

            const uploadParams = {
                Bucket: c.env.R2_BUCKET_STR,
                Key: uniqueFilename,
                Body: Buffer.from(arrayBuffer), 
                ContentType: file.type || 'image/jpeg',
            };

            await s3.send(new PutObjectCommand(uploadParams));
            urlFoto = `${c.env.R2_PUBLIC_URL}/${uniqueFilename}`;
        }

        const inputStock = stock !== undefined && stock !== '' ? parseInt(stock) : 20;
        const discount_percentage = parseFloat(body.discount_percentage) || 0;

        await pool.prepare(`
            INSERT INTO products (shop_id, name, barcode, price, cost_price, discount_percentage, category, description, image_url, stock) 
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `).bind(shopId, nama_produk, barcode, harga, cost_price, discount_percentage, kategori, deskripsi, urlFoto, inputStock).run();
        
        return c.json({
            success: true,
            message: 'Produk baru berhasil ditambahkan!',
            data: { shop_id: shopId, nama_produk, harga, kategori, deskripsi, foto: urlFoto, stock: inputStock }
        }, 201);

    } catch (error) {
        console.error("Error saat menyimpan produk:", error);
        return c.json({ success: false, message: "Gagal menyimpan produk: " + error.message }, 500);
    }
});

app.post('/api/products/:id', verifikasiAksesWarung, cekMasaAktifSub, async (c) => {
    try {
        const pool = getDbPool(c);
        const s3 = getS3Client(c);
        const productId = parseInt(c.req.param('id'));
        const body = await c.req.parseBody();

        const name = body.nama_produk || body.name;
        const price = body.harga || body.price;
        const cost_price = parseFloat(body.cost_price) || 0;
        const category = body.kategori || body.category;
        const description = body.deskripsi || body.description || '';
        const barcode = body.barcode || null;
        
        const parsedStock = body.stock !== undefined && body.stock !== null && body.stock !== '' 
            ? parseInt(body.stock, 10) 
            : 0;

        if (!name || !price || !category) {
            return c.json({ success: false, message: 'Nama, harga, dan kategori wajib diisi.' }, 400);
        }

        const { results: existingProduct } = await pool.prepare('SELECT image_url FROM products WHERE id = ?').bind(productId).all();
        if (!existingProduct || existingProduct.length === 0) {
            return c.json({ success: false, message: 'Produk tidak ditemukan.' }, 404);
        }

        let urlFoto = existingProduct[0].image_url;
        const file = body.foto_produk;

        if (file && typeof file === 'object' && file.name) {
            const fileExtension = file.name.split('.').pop();
            const uniqueFilename = `product-${Date.now()}-${crypto.randomBytes(4).toString('hex')}.${fileExtension}`;
            const arrayBuffer = await file.arrayBuffer();

            const uploadParams = {
                Bucket: c.env.R2_BUCKET_STR,
                Key: uniqueFilename,
                Body: Buffer.from(arrayBuffer), 
                ContentType: file.type || 'image/jpeg',
            };

            await s3.send(new PutObjectCommand(uploadParams));
            urlFoto = `${c.env.R2_PUBLIC_URL}/${uniqueFilename}`;
        }

        const discount_percentage = parseFloat(body.discount_percentage) || 0;
        await pool.prepare(`
            UPDATE products 
            SET name = ?, barcode = ?, price = ?, cost_price = ?, discount_percentage = ?, category = ?, description = ?, image_url = ?, stock = ?
            WHERE id = ?
        `).bind(name, barcode, parseFloat(price), cost_price, discount_percentage, category, description, urlFoto, parsedStock, productId).run();    
        
        return c.json({
            success: true,
            message: 'Produk berhasil diperbarui!',
            data: { id: productId, name, price, category, description, image_url: urlFoto, stock: parsedStock }
        });

    } catch (error) {
        console.error("Error saat memperbarui produk:", error);
        return c.json({ success: false, message: "Gagal memperbarui produk: " + error.message }, 500);
    }
});

app.get('/api/products/export-excel', verifikasiAksesWarung, async (c) => {
    try {
        const pool = getDbPool(c);
        const shopSlug = c.req.query('shop') || c.req.header('x-shop-slug');

        const shopId = await getShopIdBySlug(pool, shopSlug);
        if (!shopId) {
            return c.json({ success: false, message: 'Warung tidak ditemukan.' }, 404);
        }

        const { results: rows } = await pool.prepare(
            `SELECT barcode, name, category, cost_price, price, discount_percentage, stock, description
             FROM products 
             WHERE shop_id = ? AND is_active = 1 
             ORDER BY category ASC, name ASC`
        ).bind(shopId).all();

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Master Produk');

        worksheet.columns = [
            { header: 'Barcode', key: 'barcode', width: 18 },
            { header: 'Nama Produk', key: 'name', width: 30 },
            { header: 'Kategori', key: 'category', width: 18 },
            { header: 'HPP / Modal (Rp)', key: 'cost_price', width: 18 },
            { header: 'Harga Jual (Rp)', key: 'price', width: 18 },
            { header: 'Diskon (%)', key: 'discount_percentage', width: 12 },
            { header: 'Stok Terakhir', key: 'stock', width: 15 },
            { header: 'Deskripsi', key: 'description', width: 35 }
        ];

        worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFF' } };
        worksheet.getRow(1).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '059669' }
        };

        rows.forEach(p => {
            worksheet.addRow({
                barcode: p.barcode || '-',
                name: p.name,
                category: p.category || 'Umum',
                cost_price: parseFloat(p.cost_price) || 0,
                price: parseFloat(p.price) || 0,
                discount_percentage: parseFloat(p.discount_percentage) || 0,
                stock: parseInt(p.stock) || 0,
                description: p.description || '-'
            });
        });

        const buffer = await workbook.xlsx.writeBuffer();
        return c.body(buffer, 200, {
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Content-Disposition': `attachment; filename="Master_Produk_${shopSlug}_${Date.now()}.xlsx"`
        });
    } catch (error) {
        console.error("Error export excel master produk:", error);
        return c.json({ success: false, message: "Gagal mengeksport data master produk ke Excel." }, 500);
    }
});

app.get('/api/products/:id', verifikasiAksesWarung, async (c) => {
    try {
        const pool = getDbPool(c);
        const productId = parseInt(c.req.param('id'));
        const shopSlug = c.req.query('shop') || c.req.header('x-shop-slug');
        const shopId = await getShopIdBySlug(pool, shopSlug);

        if (!shopId) {
            return c.json({ success: false, message: "Warung tidak ditemukan." }, 404);
        }

        const { results: rows } = await pool.prepare(
            `SELECT id, name, barcode, price, cost_price, discount_percentage, category, image_url, description, is_available, stock 
            FROM products 
            WHERE id = ? AND shop_id = ? AND is_active = 1`
        ).bind(productId, shopId).all();

        if (!rows || rows.length === 0) {
            return c.json({ success: false, message: "Produk tidak ditemukan." }, 404);
        }

        return c.json({ success: true, data: rows[0] });
    } catch (error) {
        console.error("Error mengambil detail produk:", error);
        return c.json({ success: false, message: "Gagal mengambil detail produk." }, 500);
    }
});

// ---------------- KATEGORI ----------------
app.get('/api/categories', verifikasiAksesWarung, async (c) => {
    try {
        const pool = getDbPool(c);
        let shopSlug = c.req.query('shop') || c.req.header('x-shop-slug');
        let shopId = c.req.header('x-shop-id');

        if ((!shopId || shopId === 'null' || shopId === 'undefined') && shopSlug) {
            shopId = await getShopIdBySlug(pool, shopSlug);
        }

        if (!shopId) {
            return c.json({ success: false, message: "Warung tidak ditemukan.", data: [] }, 404);
        }

        const { results: rows } = await pool.prepare(
            `SELECT id, name FROM categories WHERE shop_id = ? AND is_active = 1 ORDER BY name ASC`
        ).bind(shopId).all();

        return c.json({ success: true, data: rows });
    } catch (error) {
        console.error("Error mengambil data kategori:", error);
        return c.json({ success: false, message: "Gagal mengambil data kategori: " + error.message }, 500);
    }
});

app.post('/api/categories', verifikasiAksesWarung, cekMasaAktifSub, async (c) => {
    try {
        const pool = getDbPool(c);
        const body = await c.req.json();
        const { name } = body;
        const shopSlug = c.req.query('shop') || c.req.header('x-shop-slug') || body.shop;

        if (!name || name.trim() === '') {
            return c.json({ success: false, message: "Nama kategori wajib diisi." }, 400);
        }

        const shopId = await getShopIdBySlug(pool, shopSlug);
        if (!shopId) {
            return c.json({ success: false, message: "Warung tidak ditemukan." }, 404);
        }

        await pool.prepare(
            `INSERT INTO categories (shop_id, name, is_active) VALUES (?, ?, 1)`
        ).bind(shopId, name.trim()).run();

        return c.json({ success: true, message: "Kategori berhasil ditambahkan." }, 201);
    } catch (error) {
        console.error("Error tambah kategori:", error);
        return c.json({ success: false, message: "Gagal menambahkan kategori: " + error.message }, 500);
    }
});

app.put('/api/categories/:id', verifikasiAksesWarung, cekMasaAktifSub, async (c) => {
    try {
        const pool = getDbPool(c);
        const categoryId = parseInt(c.req.param('id'));
        const body = await c.req.json();
        const { name } = body;

        if (!name || name.trim() === '') {
            return c.json({ success: false, message: "Nama kategori tidak boleh kosong." }, 400);
        }

        await pool.prepare(
            `UPDATE categories SET name = ? WHERE id = ?`
        ).bind(name.trim(), categoryId).run();

        return c.json({ success: true, message: "Kategori berhasil diperbarui." });
    } catch (error) {
        console.error("Error update kategori:", error);
        return c.json({ success: false, message: "Gagal memperbarui kategori: " + error.message }, 500);
    }
});

app.delete('/api/categories/:id', verifikasiAksesWarung, cekMasaAktifSub, async (c) => {
    try {
        const pool = getDbPool(c);
        const categoryId = parseInt(c.req.param('id'));
        await pool.prepare(`UPDATE categories SET is_active = 0 WHERE id = ?`).bind(categoryId).run();
        return c.json({ success: true, message: "Kategori berhasil dihapus." });
    } catch (error) {
        console.error("Error hapus kategori:", error);
        return c.json({ success: false, message: "Gagal menghapus kategori: " + error.message }, 500);
    }
});

// ---------------- CHECKOUT ----------------
app.post('/api/checkout', verifikasiAksesWarung, cekMasaAktifSub, async (c) => {
    const pool = getDbPool(c);
    const s3 = getS3Client(c);

    try {
        const body = await c.req.parseBody();
        const shopSlug = body.shop || c.req.query('shop') || c.req.header('x-shop-slug');
        const subtotal = body.subtotal;
        const discount = body.discount;
        const tax = body.tax;
        const total = body.total;
        const payment = body.payment;
        const payment_method = body.payment_method || 'cash';
        const customer_name = body.customer_name || body.pelanggan || null;
        
        let cart = body.cart;
        if (typeof cart === 'string') {
            cart = JSON.parse(cart);
        }

        const shopId = await getShopIdBySlug(pool, shopSlug);
        if (!shopId) {
            return c.json({ success: false, message: "Warung/Toko tidak ditemukan atau tidak valid." }, 404);
        }

        if (!cart || !Array.isArray(cart) || cart.length === 0) {
            return c.json({ success: false, message: "Keranjang belanja kosong." }, 400);
        }

        const numericSubtotal = parseFloat(subtotal) || 0;
        const numericDiscount = parseFloat(discount) || 0;
        const numericTax = parseFloat(tax) || 0;
        const numericTotal = parseFloat(total);
        const numericPayment = parseFloat(payment);

        if (isNaN(numericPayment) || numericPayment < numericTotal) {
            return c.json({ success: false, message: "Jumlah pembayaran tidak mencukupi." }, 400);
        }

        const shopCode = shopId;
        const timeBase36 = Math.floor(Date.now() / 1000).toString(36).toUpperCase();
        const randomBase36 = Math.floor(Math.random() * 1296).toString(36).toUpperCase().padStart(2, '0');
        const invoiceNumber = `INV-${shopCode}${timeBase36}-${randomBase36}`;

        let urlBuktiBayar = null;
        const proofFile = body.payment_proof;
        if (proofFile && typeof proofFile === 'object' && proofFile.name) {
            const fileExtension = proofFile.name.split('.').pop();
            const uniqueFilename = `proof-${Date.now()}-${crypto.randomBytes(4).toString('hex')}.${fileExtension}`;
            const arrayBuffer = await proofFile.arrayBuffer();

            await s3.send(new PutObjectCommand({
                Bucket: c.env.R2_BUCKET_STR,
                Key: uniqueFilename,
                Body: Buffer.from(arrayBuffer), 
                ContentType: proofFile.type || 'image/jpeg',
            }));
            urlBuktiBayar = `${c.env.R2_PUBLIC_URL}/${uniqueFilename}`;
        }

        const { results: shopRows } = await pool.prepare(
            'SELECT package_id, is_stock_calculated FROM shops WHERE id = ?'
        ).bind(shopId).all();
        
        const currentShop = shopRows[0];
        const isEligiblePackage = (currentShop.package_id === 2 || currentShop.package_id === 3);
        const checkStockActive = isEligiblePackage && (currentShop.is_stock_calculated === 1);

        if (checkStockActive) {
            const productIds = cart.map(item => item.id);
            const placeholders = productIds.map(() => '?').join(',');
            const { results: pRows } = await pool.prepare(
                `SELECT id, stock FROM products WHERE id IN (${placeholders}) AND shop_id = ?`
            ).bind(...productIds, shopId).all();

            const stockMap = new Map();
            pRows.forEach(p => stockMap.set(p.id, p.stock));

            for (const item of cart) {
                const oldStock = stockMap.get(item.id) || 0;
                const newStock = oldStock - item.qty;

                await pool.prepare('UPDATE products SET stock = ? WHERE id = ? AND shop_id = ?')
                    .bind(newStock, item.id, shopId).run();

                await pool.prepare(
                    `INSERT INTO stock_mutations (shop_id, product_id, type, qty, stock_before, stock_after, reference_number, notes) VALUES (?, ?, 'OUT', ?, ?, ?, ?, 'Penjualan POS')`
                ).bind(shopId, item.id, item.qty, oldStock, newStock, invoiceNumber).run();
            }
        }

        const changeAmount = numericPayment - numericTotal;
        const discount_percentage = parseFloat(body.discount_percentage) || 0;

        const orderResult = await pool.prepare(
            `INSERT INTO orders (invoice_number, customer_name, shop_id, payment_method, subtotal, discount, discount_percentage, tax, total, payment, \`change\`, payment_proof_url, status) 
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'completed')`
        ).bind(invoiceNumber, customer_name, shopId, payment_method, numericSubtotal, numericDiscount, discount_percentage, numericTax, numericTotal, numericPayment, changeAmount, urlBuktiBayar).run();

        const orderId = orderResult.meta.last_row_id;

        for (const item of cart) {
            await pool.prepare(
                `INSERT INTO order_details (order_id, product_id, product_name, price, cost_price, qty, subtotal) VALUES (?, ?, ?, ?, ?, ?, ?)`
            ).bind(orderId, item.id, item.name, item.price, parseFloat(item.cost_price || item.hpp) || 0, item.qty, item.price * item.qty).run();
        }

        if (!isEligiblePackage || currentShop.package_id !== 3) {
            await pool.prepare(
                'UPDATE shops SET max_transactions_monthly = MAX(0, max_transactions_monthly - 1) WHERE id = ?'
            ).bind(shopId).run();
        }

        return c.json({
            success: true,
            message: "Transaksi berhasil diproses.",
            receipt: {
                invoice_number: invoiceNumber,
                customer_name: customer_name,
                order_id: orderId,
                shop_id: shopId,
                payment_method: payment_method,
                payment_proof_url: urlBuktiBayar,
                subtotal: numericSubtotal.toLocaleString('id-ID'),
                discount: numericDiscount.toLocaleString('id-ID'),
                tax: numericTax.toLocaleString('id-ID'),
                total: numericTotal.toLocaleString('id-ID'),
                payment: numericPayment.toLocaleString('id-ID'),
                change: changeAmount.toLocaleString('id-ID'),
                items: cart
            }
        });

    } catch (error) {
        console.error("Error checkout POS:", error);
        return c.json({ success: false, message: "Gagal memproses transaksi: " + error.message }, 500);
    }
});

// ---------------- SUBSCRIPTION / LANGGANAN -----------------
app.get('/api/packages', async (c) => {
    try {
        const pool = getDbPool(c);
        const { results: rows } = await pool.prepare('SELECT * FROM packages WHERE is_active = 1 ORDER BY price_monthly ASC').all();
        return c.json({ success: true, data: rows });
    } catch (error) {
        console.error("Error ambil daftar paket:", error);
        return c.json({ success: false, message: "Gagal mengambil daftar paket langganan." }, 500);
    }
});

app.post('/api/shops/create-midtrans-qris', verifikasiAksesWarung, async (c) => {
    const pool = getDbPool(c);
    try {
        const snap = getSnapClient(c);
        const body = await c.req.json();
        const { shop, package_id, billing_cycle } = body;

        if (!package_id || !billing_cycle) {
            return c.json({ success: false, message: "Paket dan siklus tagihan wajib dipilih." }, 400);
        }

        const shopId = await getShopIdBySlug(pool, shop);
        if (!shopId) return c.json({ success: false, message: "Warung tidak ditemukan." }, 404);

        const { results: pkgRows } = await pool.prepare('SELECT name, price_monthly, price_yearly FROM packages WHERE id = ?').bind(package_id).all();
        if (!pkgRows || pkgRows.length === 0) return c.json({ success: false, message: "Paket tidak ditemukan." }, 404);

        const pkg = pkgRows[0];
        const cycle = billing_cycle === 'yearly' ? 'yearly' : 'monthly';
        const amount = cycle === 'yearly' ? pkg.price_yearly : pkg.price_monthly;
        const orderId = `SUB-${shopId}-${Date.now()}`;

        const parameter = {
            transaction_details: { 
                order_id: orderId, 
                gross_amount: Math.round(amount) 
            },
            item_details: [{
                id: `PKG-${package_id}`,
                price: Math.round(amount),
                quantity: 1,
                name: `Paket ${pkg.name} (${cycle.toUpperCase()})`
            }]
        };

        const transaction = await snap.createTransaction(parameter);
        const startDateStr = new Date().toISOString().split('T')[0];

        await pool.prepare(
            `INSERT INTO subscriptions 
            (order_id, shop_id, package_id, package_name, amount, start_date, end_date, status, payment_proof_url, billing_cycle) 
            VALUES (?, ?, ?, ?, ?, ?, ?, 'pending', ?, ?)`
        ).bind(orderId, shopId, parseInt(package_id), `${pkg.name} (${cycle.toUpperCase()})`, amount, startDateStr, startDateStr, orderId, cycle).run();

        return c.json({
            success: true,
            order_id: orderId,
            gross_amount: amount,
            snap_token: transaction.token
        });
    } catch (error) {
        console.error("Error generate Midtrans Snap:", error);
        return c.json({ success: false, message: "Gagal membuat transaksi Midtrans: " + error.message }, 500);
    }
});

app.get('/api/shops/check-midtrans-status/:orderId', verifikasiAksesWarung, async (c) => {
    try {
        const pool = getDbPool(c);
        const snap = getSnapClient(c);
        const orderId = c.req.param('orderId');

        const statusResponse = await snap.transaction.status(orderId);
        const transactionStatus = statusResponse.transaction_status;
        const fraudStatus = statusResponse.fraud_status;

        if (transactionStatus === 'settlement' || (transactionStatus === 'capture' && fraudStatus === 'accept')) {
            try {
                const { results: subRows } = await pool.prepare(
                    `SELECT id, shop_id, package_id, billing_cycle 
                     FROM subscriptions 
                     WHERE (order_id = ? OR payment_proof_url = ?) AND status = 'pending'`
                ).bind(orderId, orderId).all();

                if (subRows && subRows.length > 0) {
                    const sub = subRows[0];
                    const daysToAdd = sub.billing_cycle === 'yearly' ? 365 : 30;

                    const { results: shopRows } = await pool.prepare(
                        `SELECT subscription_until, max_transactions_monthly, package_id FROM shops WHERE id = ?`
                    ).bind(sub.shop_id).all();

                    const shop = shopRows[0];
                    const hariIni = new Date();

                    const { results: pkgRows } = await pool.prepare(
                        'SELECT id, max_transactions_monthly FROM packages WHERE id = ?'
                    ).bind(sub.package_id).all();
                    const pkgMaxTx = pkgRows.length > 0 ? pkgRows[0].max_transactions_monthly : 0;
                    const isNewSultan = (sub.package_id === 3);

                    let newUntilDate = new Date();
                    let newQuota = 0;

                    if (shop && shop.subscription_until && new Date(shop.subscription_until) > hariIni) {
                        const baseDate = new Date(shop.subscription_until);
                        baseDate.setDate(baseDate.getDate() + daysToAdd);
                        newUntilDate = baseDate;

                        if (isNewSultan) {
                            newQuota = 0;
                        } else {
                            const currentQuota = Math.max(0, parseInt(shop.max_transactions_monthly) || 0);
                            newQuota = currentQuota + pkgMaxTx;
                        }
                    } else {
                        const baseDate = new Date();
                        baseDate.setDate(baseDate.getDate() + daysToAdd);
                        newUntilDate = baseDate;

                        newQuota = isNewSultan ? 0 : pkgMaxTx;
                    }

                    const startDateStr = hariIni.toISOString().split('T')[0];
                    const newUntilStr = newUntilDate.toISOString().split('T')[0];

                    await pool.prepare(
                        `UPDATE shops 
                        SET subscription_status = 'active', 
                            subscription_until = ?, 
                            package_id = ?, 
                            billing_cycle = ?,
                            max_transactions_monthly = ? 
                        WHERE id = ?`
                    ).bind(newUntilStr, sub.package_id, sub.billing_cycle, newQuota, sub.shop_id).run();

                    await pool.prepare(
                        `UPDATE subscriptions 
                        SET status = 'active', 
                            start_date = ?, 
                            end_date = ?,
                            max_transactions_monthly = ?
                        WHERE id = ?`
                    ).bind(startDateStr, newUntilStr, newQuota, sub.id).run();
                }
            } catch (err) {
                console.error("Error update DB via Polling status:", err);
            }
        }

        return c.json({
            success: true,
            order_id: orderId,
            transaction_status: transactionStatus,
            fraud_status: fraudStatus
        });
    } catch (error) {
        console.error("Gagal cek status Midtrans:", error);
        return c.json({ 
            success: false, 
            message: "Gagal memeriksa status pembayaran Midtrans." 
        }, 500);
    }
});

app.get('/api/shops/subscription', verifikasiAksesWarung, async (c) => {
    try {
        const pool = getDbPool(c);
        const shopSlug = c.req.query('shop');
        const shopId = await getShopIdBySlug(pool, shopSlug);
        if (!shopId) return c.json({ success: false, message: "Warung tidak ditemukan." }, 404);

        const { results: rows } = await pool.prepare(
            `SELECT s.id, s.package_id, s.subscription_status, s.subscription_until, s.max_transactions_monthly,
                    p.name AS package_name, p.slug AS package_slug
             FROM shops s
             LEFT JOIN packages p ON s.package_id = p.id
             WHERE s.id = ?`
        ).bind(shopId).all();

        if (!rows || rows.length === 0) return c.json({ success: false, message: "Data toko tidak ditemukan." }, 404);

        const shop = rows[0];
        const sekarang = new Date();
        const subUntil = shop.subscription_until ? new Date(shop.subscription_until) : null;
        const packageId = Number(shop.package_id || 1);
        let remainingQuota = shop.max_transactions_monthly || 0;

        let remainingDays = 0;
        let isExpiredByDate = false;
        let isToleransi = false;

        if (subUntil) {
            const diffTime = subUntil - sekarang;
            remainingDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
            const batasToleransi = new Date(subUntil);
            batasToleransi.setDate(batasToleransi.getDate() + 1);

            if (sekarang > batasToleransi) {
                isExpiredByDate = true;
            } else if (sekarang > subUntil && sekarang <= batasToleransi) {
                isToleransi = true;
            }
        }

        if (isExpiredByDate && shop.max_transactions_monthly > 0) {
            await pool.prepare('UPDATE shops SET max_transactions_monthly = 0 WHERE id = ?').bind(shopId).run();
            remainingQuota = 0;
        }

        const isSultan = (packageId === 3 || (shop.package_slug && shop.package_slug.toLowerCase() === 'sultan'));
        let isExpired = isSultan ? isExpiredByDate : (isExpiredByDate || remainingQuota <= 0);

        return c.json({
            success: true,
            data: {
                package_id: packageId,
                package_name: shop.package_name || 'UMKM',
                subscription_status: shop.subscription_status,
                remaining_days: remainingDays,
                is_expired: isExpired,
                is_toleransi: isToleransi,
                is_trial: shop.subscription_status === 'trial',
                remaining_transactions: remainingQuota,
                quota_text: isSultan ? "Kuota Unlimited" : `Sisa Kuota: ${remainingQuota} Transaksi`
            }
        });
    } catch (error) {
        console.error("Error cek subscription:", error);
        return c.json({ success: false, message: "Gagal mengambil data langganan." }, 500);
    }
});

app.post('/api/shops/subscribe', verifikasiAksesWarung, async (c) => {
    const pool = getDbPool(c);
    const s3 = getS3Client(c);
    const body = await c.req.parseBody();
    const { shop, package_id, billing_cycle } = body;
    const proofFile = body.payment_proof_url;

    if (!package_id || !proofFile) {
        return c.json({ success: false, message: "Paket langganan dan bukti pembayaran wajib disertakan." }, 400);
    }

    try {
        const shopId = await getShopIdBySlug(pool, shop);
        if (!shopId) {
            return c.json({ success: false, message: "Warung tidak ditemukan." }, 404);
        }

        const selectedPackageId = parseInt(package_id);
        const { results: pkgRows } = await pool.prepare('SELECT name, price_monthly, price_yearly FROM packages WHERE id = ?').bind(selectedPackageId).all();
        if (!pkgRows || pkgRows.length === 0) {
            return c.json({ success: false, message: "Paket tidak ditemukan." }, 404);
        }

        const packageData = pkgRows[0];
        const cycle = (billing_cycle === 'yearly') ? 'yearly' : 'monthly';
        const nominalBayar = cycle === 'yearly' ? packageData.price_yearly : packageData.price_monthly;

        const fileExtension = proofFile.name.split('.').pop();
        const uniqueFilename = `sub-proof-${Date.now()}-${crypto.randomBytes(4).toString('hex')}.${fileExtension}`;
        const arrayBuffer = await proofFile.arrayBuffer();

        await s3.send(new PutObjectCommand({
            Bucket: c.env.R2_BUCKET_STR,
            Key: uniqueFilename,
            Body: Buffer.from(arrayBuffer),
            ContentType: proofFile.type || 'image/jpeg',
        }));

        const urlBuktiBayar = `${c.env.R2_PUBLIC_URL}/${uniqueFilename}`;
        const startDateStr = new Date().toISOString().split('T')[0];

        await pool.prepare(
            `UPDATE shops SET subscription_status = 'pending', package_id = ?, billing_cycle = ? WHERE id = ?`
        ).bind(selectedPackageId, cycle, shopId).run();

        await pool.prepare(
            `INSERT INTO subscriptions 
            (shop_id, package_id, package_name, amount, start_date, end_date, status, payment_proof_url, billing_cycle) 
            VALUES (?, ?, ?, ?, ?, ?, 'pending', ?, ?)`
        ).bind(shopId, selectedPackageId, `${packageData.name} (${cycle.toUpperCase()})`, nominalBayar, startDateStr, startDateStr, urlBuktiBayar, cycle).run();

        return c.json({ success: true, message: "Bukti pembayaran berhasil dikirim!" });
    } catch (error) {
        console.error("Error submit perpanjangan paket:", error);
        return c.json({ success: false, message: "Gagal memproses bukti pembayaran: " + error.message }, 500);
    }
});

// ---------------- AUTH, DASHBOARD, REGISTER ----------------
app.post('/api/auth/login', async (c) => {
    try {
        const db = getDbPool(c); 
        const { username, password } = await c.req.json();
        
        if (!username || !password) {
            return c.json({ success: false, message: 'Username dan password wajib diisi!' }, 400);
        }

        const { results } = await db
            .prepare('SELECT * FROM shops WHERE username = ? AND password = ?')
            .bind(username, password)
            .all();

        if (!results || results.length === 0) {
            return c.json({ success: false, message: 'Username atau password salah.' }, 401);
        }

        const shop = results[0];
        return c.json({
            success: true,
            message: 'Login berhasil!',
            shop_id: shop.id,
            shop_name: shop.shop_name,
            slug: shop.slug,
            app_logo: c.env.APP_LOGO_URL || "https://pub-c3b5b9a8f041497f97f050b2133dbd3a.r2.dev/logo.png"
        });
    } catch (error) {
        console.error("Error saat login:", error);
        return c.json({ success: false, message: 'Terjadi kesalahan internal server: ' + error.message }, 500);
    }
});

app.get('/api/orders/dashboard-pos', verifikasiAksesWarung, async (c) => {
    try {
        const pool = getDbPool(c);
        const shopSlug = c.req.query('shop') || c.req.header('x-shop-slug');
        const startDate = c.req.query('startDate');
        const endDate = c.req.query('endDate');

        const shopId = await getShopIdBySlug(pool, shopSlug);
        if (!shopId) {
            return c.json({ success: false, message: "Warung tidak ditemukan." }, 404);
        }

        let queryText = `
            SELECT o.id, o.invoice_number, o.customer_name, o.payment_method, o.payment_proof_url, 
                   o.subtotal, o.discount, o.tax, o.total, o.payment, o.\`change\`, o.status, 
                   o.cancelled_at, o.cancel_reason, o.created_at, s.wilayah 
            FROM orders o
            JOIN shops s ON o.shop_id = s.id
            WHERE o.shop_id = ?
        `;
        let params = [shopId];

        if (startDate && endDate) {
            queryText += ` AND DATE(o.created_at) BETWEEN ? AND ?`;
            params.push(startDate, endDate);
        } else {
            queryText += ` AND DATE(o.created_at) = DATE('now')`;
        }

        queryText += ` ORDER BY o.created_at DESC`;

        const { results: rows } = await pool.prepare(queryText).bind(...params).all();
        return c.json({ success: true, orders: rows });
    } catch (error) {
        console.error("Error dashboard POS:", error);
        return c.json({ success: false, message: "Gagal mengambil data transaksi dashboard." }, 500);
    }
});

app.post('/api/orders/cancel/:orderId', verifikasiAksesWarung, cekMasaAktifSub, async (c) => {
    const pool = getDbPool(c);
    try {
        const orderId = c.req.param('orderId');
        const body = await c.req.json();
        const shopSlug = body.shop || c.req.query('shop') || c.req.header('x-shop-slug');
        const cancelReason = body.cancel_reason || body.reason || 'Salah input / Batal transaksi';

        const shopId = await getShopIdBySlug(pool, shopSlug);
        if (!shopId) {
            return c.json({ success: false, message: "Warung tidak ditemukan." }, 404);
        }

        const { results: orders } = await pool.prepare(
            'SELECT id, status, invoice_number FROM orders WHERE id = ? AND shop_id = ?'
        ).bind(orderId, shopId).all();

        if (!orders || orders.length === 0) {
            return c.json({ success: false, message: "Transaksi tidak ditemukan." }, 404);
        }

        if (orders[0].status === 'cancelled') {
            return c.json({ success: false, message: "Transaksi ini sudah dibatalkan sebelumnya." }, 400);
        }

        const { results: items } = await pool.prepare(
            'SELECT product_id, qty FROM order_details WHERE order_id = ?'
        ).bind(orderId).all();

        for (const item of items) {
            if (item.product_id) {
                const { results: pRows } = await pool.prepare(
                    'SELECT stock FROM products WHERE id = ? AND shop_id = ?'
                ).bind(item.product_id, shopId).all();
                const oldStock = pRows[0] ? pRows[0].stock : 0;
                const newStock = oldStock + item.qty;

                await pool.prepare(
                    'UPDATE products SET stock = ? WHERE id = ?'
                ).bind(newStock, item.product_id).run();

                await pool.prepare(
                    `INSERT INTO stock_mutations (shop_id, product_id, type, qty, stock_before, stock_after, reference_number, notes) 
                    VALUES (?, ?, 'IN', ?, ?, ?, ?, ?)`
                ).bind(shopId, item.product_id, item.qty, oldStock, newStock, orders[0].invoice_number || `CANCEL-${orderId}`, `Batal Tx: ${cancelReason}`).run();
            }
        }

        await pool.prepare(
            "UPDATE orders SET status = 'cancelled', cancelled_at = CURRENT_TIMESTAMP, cancel_reason = ? WHERE id = ?"
        ).bind(cancelReason, orderId).run();

        return c.json({ success: true, message: "Transaksi berhasil dibatalkan dan stok produk telah dikembalikan." });

    } catch (error) {
        console.error("Error batalkan transaksi:", error);
        return c.json({ success: false, message: "Gagal membatalkan transaksi: " + error.message }, 500);
    }
});

app.get('/api/shops/settings', verifikasiAksesWarung, async (c) => {
    try {
        const pool = getDbPool(c);
        const shopSlug = c.req.query('shop') || c.req.header('x-shop-slug');
        const { results: rows } = await pool.prepare(
            `SELECT id, shop_name, owner_name, slug, is_open, show_cash_payment, has_tax, 
                    tax_percentage, discount_percentage, bank_rekening_info, qris_image_url,
                    is_stock_calculated, package_id 
             FROM shops WHERE slug = ?`
        ).bind(shopSlug).all();

        if (!rows || rows.length === 0) {
            return c.json({ success: false, message: "Warung tidak ditemukan." }, 404);
        }

        return c.json({ success: true, data: rows[0] });
    } catch (error) {
        console.error("Error ambil setting toko:", error);
        return c.json({ success: false, message: "Gagal mengambil data setting toko." }, 500);
    }
});

app.put('/api/shops/settings', verifikasiAksesWarung, cekMasaAktifSub, async (c) => {
    try {
        const pool = getDbPool(c);
        const s3 = getS3Client(c);
        const body = await c.req.parseBody();

        const shopSlug = body.shop || c.req.query('shop') || c.req.header('x-shop-slug');
        const shopId = await getShopIdBySlug(pool, shopSlug);
        if (!shopId) {
            return c.json({ success: false, message: "Warung tidak ditemukan." }, 404);
        }

        const show_cash_payment = body.show_cash_payment === '1' || body.show_cash_payment === 1 ? 1 : 0;
        const has_tax = body.has_tax === '1' || body.has_tax === 1 ? 1 : 0;
        const is_stock_calculated = body.is_stock_calculated === '1' || body.is_stock_calculated === 1 ? 1 : 0;
        const tax_percentage = parseFloat(body.tax_percentage) || 0;
        const discount_percentage = parseFloat(body.discount_percentage) || 0;
        const bank_rekening_info = body.bank_rekening_info || null;

        let qrisUrlQuery = "";
        let params = [show_cash_payment, has_tax, is_stock_calculated, tax_percentage, discount_percentage, bank_rekening_info];

        const file = body.qris_image;
        if (file && typeof file === 'object' && file.name) {
            const fileExtension = file.name.split('.').pop().toLowerCase();
            const uniqueFilename = `qris-${Date.now()}-${crypto.randomBytes(4).toString('hex')}.${fileExtension}`;
            const arrayBuffer = await file.arrayBuffer();
            const binaryData = new Uint8Array(arrayBuffer);

            let mimeType = file.type || (fileExtension === 'png' ? 'image/png' : 'image/jpeg');

            // Simpan ke R2 via binding Worker
            await c.env.R2_BUCKET.put(uniqueFilename, binaryData, {
                httpMetadata: { contentType: mimeType }
            });

            // Simpan rute API internal ke Database
            const urlFoto = `/api/images/${uniqueFilename}`;
            qrisUrlQuery = ", qris_image_url = ?";
            params.push(urlFoto);
        }

        params.push(shopId);

        const queryText = `
            UPDATE shops 
            SET show_cash_payment = ?, has_tax = ?, is_stock_calculated = ?, tax_percentage = ?, discount_percentage = ?, bank_rekening_info = ? ${qrisUrlQuery} 
            WHERE id = ?
        `;

        await pool.prepare(queryText).bind(...params).run();

        return c.json({ success: true, message: "Pengaturan toko berhasil diperbarui!" });
    } catch (error) {
        console.error("Error update setting toko:", error);
        return c.json({ success: false, message: "Gagal menyimpan pengaturan toko: " + error.message }, 500);
    }
});
// ---------------- IMAGE PROXY SERVER (UNTUK SEMUA GAMBAR R2) ----------------
app.get('/api/images/:key', async (c) => {
    try {
        const key = c.req.param('key');
        if (!key) return c.text('Key gambar tidak ditemukan', 400);

        // Ambil objek gambar langsung dari R2 via internal binding
        const object = await c.env.R2_BUCKET.get(key);
        if (!object) {
            return c.text('Gambar tidak ditemukan di R2', 404);
        }

        const headers = new Headers();
        object.writeHttpMetadata(headers);
        headers.set('etag', object.httpEtag);
        headers.set('Access-Control-Allow-Origin', '*'); // Bebas CORS di semua browser
        headers.set('Cache-Control', 'public, max-age=31536000'); // Cache gambar 1 tahun

        return new Response(object.body, { headers });
    } catch (error) {
        console.error("Error serving image via Worker:", error);
        return c.text("Gagal memuat gambar: " + error.message, 500);
    }
});

app.get('/api/orders/details/:orderId', verifikasiAksesWarung, async (c) => {
    try {
        const pool = getDbPool(c);
        const orderId = c.req.param('orderId');
        const shopSlug = c.req.query('shop') || c.req.header('x-shop-slug');
        const shopId = await getShopIdBySlug(pool, shopSlug);

        const { results: orderRows } = await pool.prepare(
            `SELECT id, invoice_number, customer_name, payment_method, subtotal, discount, tax, total, payment, \`change\`, created_at 
             FROM orders 
             WHERE id = ? AND shop_id = ?`
        ).bind(orderId, shopId).all();

        if (!orderRows || orderRows.length === 0) {
            return c.json({ success: false, message: "Data pesanan tidak ditemukan." }, 404);
        }

        const { results: itemRows } = await pool.prepare(
            `SELECT product_name, price, qty, subtotal FROM order_details WHERE order_id = ?`
        ).bind(orderId).all();

        return c.json({ 
            success: true, 
            order: orderRows[0], 
            items: itemRows 
        });
    } catch (error) {
        console.error("Error order details:", error);
        return c.json({ success: false, message: "Gagal mengambil detail item order." }, 500);
    }
});

app.post('/api/register', async (c) => {
    const pool = getDbPool(c);
    const body = await c.req.json();
    const { owner_name, shop_name, slug, username, password, package_id, billing_cycle, wilayah, is_terms_agreed } = body;

    if (!owner_name || !shop_name || !slug || !username || !password) {
        return c.json({ success: false, message: "Semua field wajib diisi." }, 400);
    }

    try {
        const { results: existingShop } = await pool.prepare(
            "SELECT id FROM shops WHERE slug = ? OR username = ?"
        ).bind(slug, username).all();

        if (existingShop && existingShop.length > 0) {
            return c.json({ 
                success: false, 
                message: "Nama warung (slug) atau Nomor HP (username) sudah terdaftar!" 
            }, 400);
        }

        const selectedPackageId = package_id ? parseInt(package_id) : 1;
        const { results: pkgRows } = await pool.prepare(
            'SELECT name, price_monthly, price_yearly, max_transactions_monthly FROM packages WHERE id = ?'
        ).bind(selectedPackageId).all();
        const packageData = pkgRows && pkgRows.length > 0 ? pkgRows[0] : { name: 'UMKM', max_transactions_monthly: 0 };
        const cycle = (billing_cycle === 'yearly') ? 'yearly' : 'monthly';

        const startDate = new Date();
        const endDate = new Date();
        endDate.setDate(startDate.getDate() + 14);

        const startDateStr = startDate.toISOString().split('T')[0];
        const endDateStr = endDate.toISOString().split('T')[0];
        const selectedWilayah = wilayah || 'DKI Jakarta';
        const termsStatus = is_terms_agreed !== undefined ? Number(is_terms_agreed) : 1;

        const shopResult = await pool.prepare(
            `INSERT INTO shops 
            (shop_name, owner_name, slug, username, password, is_open, wilayah, subscription_status, subscription_until, package_id, billing_cycle, max_transactions_monthly, is_terms_agreed) 
            VALUES (?, ?, ?, ?, ?, 1, ?, 'trial', ?, ?, ?, ?, ?)`
        ).bind(shop_name, owner_name, slug, username, password, selectedWilayah, endDateStr, selectedPackageId, cycle, 50, termsStatus).run();

        const newShopId = shopResult.meta.last_row_id;

        await pool.prepare(
            `INSERT INTO subscriptions 
            (shop_id, package_id, package_name, amount, start_date, end_date, status, billing_cycle, max_transactions_monthly) 
            VALUES (?, ?, ?, 0.00, ?, ?, 'active', ?, ?)`
        ).bind(newShopId, selectedPackageId, `Trial 14 Hari (${packageData.name} - ${cycle.toUpperCase()})`, startDateStr, endDateStr, cycle, 50).run();

        const smsMessage = `Kami dari BEDApos, ${shop_name} (TRIAL 14hr), Link Aplikasi: pos.bedadigital.app/login.html `;
   
        await pool.prepare(
            `INSERT INTO sms_queue (phone, message, status, retry_count) VALUES (?, ?, 'PENDING', 0)`
        ).bind(username, smsMessage).run();

        return c.json({
            success: true,
            message: "Registrasi toko berhasil!",
            shop_id: newShopId,
            slug: slug,
            subscription_until: endDateStr
        }, 201);

    } catch (error) {
        console.error("Error pendaftaran warung:", error);
        return c.json({ success: false, message: "Gagal menyimpan data pendaftaran: " + error.message }, 500);
    }
});

app.post('/api/auth/reset-password', async (c) => {
    try {
        const pool = getDbPool(c);
        const { username, new_password, shop } = await c.req.json();

        if (!username || !new_password) {
            return c.json({ success: false, message: 'No. Handphone (username) dan password baru wajib diisi!' }, 400);
        }

        let query = 'SELECT id FROM shops WHERE username = ?';
        let params = [username];

        if (shop) {
            query += ' AND slug = ?';
            params.push(shop);
        }

        const { results: rows } = await pool.prepare(query).bind(...params).all();

        if (!rows || rows.length === 0) {
            return c.json({ 
                success: false, 
                message: 'Data warung dengan nomor handphone tersebut tidak ditemukan!' 
            }, 404);
        }

        const shopId = rows[0].id;
        await pool.prepare('UPDATE shops SET password = ? WHERE id = ?').bind(new_password, shopId).run();

        return c.json({
            success: true,
            message: 'Password berhasil diperbarui!'
        });

    } catch (error) {
        console.error("Error saat reset password:", error);
        return c.json({ success: false, message: 'Terjadi kesalahan internal server: ' + error.message }, 500);
    }
});

// ---------------- STOK & MUTASI STOK ----------------
app.get('/api/stock-mutations', verifikasiAksesWarung, async (c) => {
    try {
        const pool = getDbPool(c);
        const shopSlug = c.req.query('shop') || c.req.header('x-shop-slug');
        const startDate = c.req.query('startDate');
        const endDate = c.req.query('endDate');

        const shopId = await getShopIdBySlug(pool, shopSlug);
        if (!shopId) {
            return c.json({ success: false, message: 'Warung tidak ditemukan.' }, 404);
        }

        const { results: invSummary } = await pool.prepare(
            `SELECT SUM(stock * cost_price) AS total_inventory_value, SUM(stock) AS total_items 
             FROM products WHERE shop_id = ? AND is_active = 1`
        ).bind(shopId).all();

        let queryText = `
            SELECT sm.id, sm.type, sm.qty, sm.buy_price, sm.unit_cost, sm.stock_before, sm.stock_after, sm.reference_number, sm.notes, sm.created_at,
                   p.name AS product_name, p.barcode
            FROM stock_mutations sm
            JOIN products p ON sm.product_id = p.id
            WHERE sm.shop_id = ?
        `;
        let params = [shopId];

        if (startDate && endDate) {
            queryText += ` AND DATE(sm.created_at) BETWEEN ? AND ?`;
            params.push(startDate, endDate);
        }

        queryText += ` ORDER BY sm.created_at DESC`;

        const { results: rows } = await pool.prepare(queryText).bind(...params).all();
        return c.json({ 
            success: true, 
            inventory_summary: invSummary[0],
            mutations: rows 
        });
    } catch (error) {
        console.error('Error laporan mutasi stok:', error);
        return c.json({ success: false, message: 'Gagal mengambil laporan mutasi stok.' }, 500);
    }
});

app.post('/api/payments/midtrans-notification', async (c) => {
    try {
        const pool = getDbPool(c);
        const snap = getSnapClient(c);
        const notification = await c.req.json();
        
        if (!notification || Object.keys(notification).length === 0) {
            return c.json({ success: true, message: "Notification test received." });
        }

        const orderId = notification.order_id || '';

        if (orderId.startsWith('BEDAORDER-') || orderId.startsWith('ORDER-')) {
            try {
                const response = await fetch('https://nodejs-pesan-antar-production.up.railway.app/api/payments/midtrans-notification', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(notification)
                });
                const resData = await response.json();
                return c.json(resData, response.status);
            } catch (fwdError) {
                return c.json({ success: true, message: "Forwarding failed but acknowledged." });
            }
        }

        if (orderId.startsWith('payment_notif_test')) {
            return c.json({ success: true, message: "Test notification standard processed." });
        }

        const statusResponse = await snap.transaction.notification(notification);
        const transactionStatus = statusResponse.transaction_status;
        const fraudStatus = statusResponse.fraud_status;

        if (transactionStatus === 'settlement' || (transactionStatus === 'capture' && fraudStatus === 'accept')) {
            try {
                const { results: subRows } = await pool.prepare(
                    `SELECT id, shop_id, package_id, billing_cycle 
                     FROM subscriptions 
                     WHERE (order_id = ? OR payment_proof_url = ?) AND status = 'pending'`
                ).bind(orderId, orderId).all();

                if (subRows && subRows.length > 0) {
                    const sub = subRows[0];
                    const daysToAdd = sub.billing_cycle === 'yearly' ? 365 : 30;

                    const { results: shopRows } = await pool.prepare(
                        `SELECT subscription_until, max_transactions_monthly FROM shops WHERE id = ?`
                    ).bind(sub.shop_id).all();

                    const { results: pkgRows } = await pool.prepare(
                        `SELECT id, max_transactions_monthly FROM packages WHERE id = ?`
                    ).bind(sub.package_id).all();

                    const pkgMaxTx = pkgRows.length > 0 ? pkgRows[0].max_transactions_monthly : 0;
                    const isNewSultan = (sub.package_id === 3);
                    const shop = shopRows[0];
                    const hariIni = new Date();

                    let newUntilDate = new Date();
                    let newQuota = 0;

                    if (shop && shop.subscription_until && new Date(shop.subscription_until) > hariIni) {
                        const baseDate = new Date(shop.subscription_until);
                        baseDate.setDate(baseDate.getDate() + daysToAdd);
                        newUntilDate = baseDate;

                        if (isNewSultan) {
                            newQuota = 0;
                        } else {
                            const currentQuota = Math.max(0, parseInt(shop.max_transactions_monthly) || 0);
                            newQuota = currentQuota + pkgMaxTx;
                        }
                    } else {
                        const baseDate = new Date();
                        baseDate.setDate(baseDate.getDate() + daysToAdd);
                        newUntilDate = baseDate;

                        newQuota = isNewSultan ? 0 : pkgMaxTx;
                    }

                    const startDateStr = hariIni.toISOString().split('T')[0];
                    const newUntilStr = newUntilDate.toISOString().split('T')[0];

                    await pool.prepare(
                        `UPDATE shops 
                        SET subscription_status = 'active', 
                            subscription_until = ?, 
                            package_id = ?, 
                            billing_cycle = ?,
                            max_transactions_monthly = ? 
                        WHERE id = ?`
                    ).bind(newUntilStr, sub.package_id, sub.billing_cycle, newQuota, sub.shop_id).run();

                    await pool.prepare(
                        `UPDATE subscriptions 
                        SET status = 'active', 
                            start_date = ?, 
                            end_date = ?,
                            max_transactions_monthly = ?
                        WHERE id = ?`
                    ).bind(startDateStr, newUntilStr, newQuota, sub.id).run();
                }
            } catch (err) {
                console.error("Gagal update DB saat notification:", err.message);
            }
        }

        return c.json({ success: true });
    } catch (error) {
        console.error("Error Webhook Midtrans:", error);
        return c.json({ success: false, message: error.message });
    }
});

// ---------------- EXPORT EXCEL RIWAYAT TRANSAKSI ----------------
app.get('/api/orders/export-excel', verifikasiAksesWarung, async (c) => {
    try {
        const pool = getDbPool(c);
        const shopSlug = c.req.query('shop') || c.req.header('x-shop-slug');
        const startDate = c.req.query('startDate');
        const endDate = c.req.query('endDate');

        const shopId = await getShopIdBySlug(pool, shopSlug);
        if (!shopId) {
            return c.json({ success: false, message: "Warung tidak ditemukan." }, 404);
        }

        const { results: shopRows } = await pool.prepare('SELECT wilayah FROM shops WHERE id = ?').bind(shopId).all();
        const shopWilayah = shopRows[0]?.wilayah || 'DKI Jakarta';
        const targetTimeZone = getTimeZoneByWilayah(shopWilayah);

        let queryText = `
            SELECT o.invoice_number, o.customer_name, o.payment_method, o.subtotal, o.discount, o.tax, o.total, o.payment, o.\`change\`, o.status, o.cancel_reason, o.created_at,
                   GROUP_CONCAT(CONCAT(od.product_name, ' (x', od.qty, ')') , ', ') AS items_detail
            FROM orders o
            LEFT JOIN order_details od ON o.id = od.order_id
            WHERE o.shop_id = ?
        `;
        let params = [shopId];

        if (startDate && endDate) {
            queryText += ` AND DATE(o.created_at) BETWEEN ? AND ?`;
            params.push(startDate, endDate);
        } else {
            queryText += ` AND DATE(o.created_at) = DATE('now')`;
        }

        queryText += ` GROUP BY o.id ORDER BY o.created_at DESC`;

        const { results: rows } = await pool.prepare(queryText).bind(...params).all();

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Riwayat Transaksi');

        worksheet.columns = [
            { header: 'No. Faktur', key: 'invoice_number', width: 22 },
            { header: 'Pelanggan', key: 'customer_name', width: 18 },
            { header: 'Tanggal & Jam', key: 'created_at', width: 22 },
            { header: 'Item / Produk Dibeli', key: 'items_detail', width: 40 },
            { header: 'Metode Bayar', key: 'payment_method', width: 15 },
            { header: 'Subtotal (Rp)', key: 'subtotal', width: 15 },
            { header: 'Diskon (Rp)', key: 'discount', width: 12 },
            { header: 'Pajak (Rp)', key: 'tax', width: 12 },
            { header: 'Total (Rp)', key: 'total', width: 15 },
            { header: 'Bayar (Rp)', key: 'payment', width: 15 },
            { header: 'Kembalian (Rp)', key: 'change', width: 15 },
            { header: 'Status', key: 'status', width: 12 },
            { header: 'Alasan Batal', key: 'cancel_reason', width: 25 }
        ];

        worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFF' } };
        worksheet.getRow(1).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '4F46E5' }
        };

        rows.forEach(r => {
            const formattedDate = new Date(r.created_at).toLocaleString('id-ID', {
                timeZone: targetTimeZone,
                day: '2-digit',
                month: '2-digit',
                year: 'numeric',
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit'
            }).replace(/\./g, ':');

            worksheet.addRow({
                invoice_number: r.invoice_number,
                customer_name: r.customer_name || '-',
                created_at: formattedDate,
                items_detail: r.items_detail || '-',
                payment_method: (r.payment_method || 'CASH').toUpperCase(),
                subtotal: parseFloat(r.subtotal) || 0,
                discount: parseFloat(r.discount) || 0,
                tax: parseFloat(r.tax) || 0,
                total: parseFloat(r.total) || 0,
                payment: parseFloat(r.payment) || 0,
                change: parseFloat(r.change) || 0,
                status: r.status === 'cancelled' ? 'BATAL' : 'SUCCESS',
                cancel_reason: r.cancel_reason || '-'
            });
        });

        const buffer = await workbook.xlsx.writeBuffer();
        return c.body(buffer, 200, {
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Content-Disposition': `attachment; filename="Riwayat_Transaksi_${shopSlug}_${Date.now()}.xlsx"`
        });
    } catch (error) {
        console.error("Error export excel transaksi:", error);
        return c.json({ success: false, message: "Gagal mengeksport data transaksi ke Excel." }, 500);
    }
});

// ---------------- EXPORT EXCEL MUTASI STOK ----------------
app.get('/api/stock-mutations/export-excel', verifikasiAksesWarung, async (c) => {
    try {
        const pool = getDbPool(c);
        const shopSlug = c.req.query('shop') || c.req.header('x-shop-slug');
        const startDate = c.req.query('startDate');
        const endDate = c.req.query('endDate');

        const shopId = await getShopIdBySlug(pool, shopSlug);
        if (!shopId) {
            return c.json({ success: false, message: 'Warung tidak ditemukan.' }, 404);
        }

        const { results: shopRows } = await pool.prepare('SELECT wilayah FROM shops WHERE id = ?').bind(shopId).all();
        const shopWilayah = shopRows[0]?.wilayah || 'DKI Jakarta';
        const targetTimeZone = getTimeZoneByWilayah(shopWilayah);

        let queryText = `
            SELECT sm.type, sm.qty, sm.buy_price, sm.unit_cost, sm.stock_before, sm.stock_after, sm.reference_number, sm.notes, sm.created_at,
                   p.name AS product_name, p.barcode
            FROM stock_mutations sm
            JOIN products p ON sm.product_id = p.id
            WHERE sm.shop_id = ?
        `;
        let params = [shopId];

        if (startDate && endDate) {
            queryText += ` AND DATE(sm.created_at) BETWEEN ? AND ?`;
            params.push(startDate, endDate);
        }

        queryText += ` ORDER BY sm.created_at DESC`;

        const { results: rows } = await pool.prepare(queryText).bind(...params).all();

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Laporan Mutasi Stok');

        worksheet.columns = [
            { header: 'Waktu', key: 'created_at', width: 22 },
            { header: 'Nama Produk', key: 'product_name', width: 25 },
            { header: 'Barcode', key: 'barcode', width: 15 },
            { header: 'Tipe Mutasi', key: 'type', width: 12 },
            { header: 'Qty', key: 'qty', width: 10 },
            { header: 'Stok Awal', key: 'stock_before', width: 12 },
            { header: 'Stok Akhir', key: 'stock_after', width: 12 },
            { header: 'Harga Beli/Pcs (Rp)', key: 'buy_price', width: 18 },
            { header: 'HPP Efektif (Rp)', key: 'unit_cost', width: 18 },
            { header: 'No. Referensi', key: 'reference_number', width: 22 },
            { header: 'Catatan', key: 'notes', width: 25 }
        ];

        worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFF' } };
        worksheet.getRow(1).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '7E22CE' }
        };

        rows.forEach(m => {
            const formattedDate = new Date(m.created_at).toLocaleString('id-ID', {
                timeZone: targetTimeZone,
                day: '2-digit',
                month: '2-digit',
                year: 'numeric',
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit'
            }).replace(/\./g, ':');

            worksheet.addRow({
                created_at: formattedDate,
                product_name: m.product_name,
                barcode: m.barcode || '-',
                type: m.type === 'IN' ? 'MASUK (+)' : 'KELUAR (-)',
                qty: m.qty,
                stock_before: m.stock_before,
                stock_after: m.stock_after,
                buy_price: parseFloat(m.buy_price) || 0,
                unit_cost: parseFloat(m.unit_cost) || 0,
                reference_number: m.reference_number || '-',
                notes: m.notes || '-'
            });
        });

        const buffer = await workbook.xlsx.writeBuffer();
        return c.body(buffer, 200, {
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Content-Disposition': `attachment; filename="Mutasi_Stok_${shopSlug}_${Date.now()}.xlsx"`
        });
    } catch (error) {
        console.error("Error export excel mutasi stok:", error);
        return c.json({ success: false, message: "Gagal mengeksport data mutasi stok ke Excel." }, 500);
    }
});

// ---------------- LAPORAN LABA RUGI ----------------
app.get('/api/reports/profit-loss', verifikasiAksesWarung, async (c) => {
    try {
        const pool = getDbPool(c);
        const shopSlug = c.req.query('shop') || c.req.header('x-shop-slug');
        const startDate = c.req.query('startDate');
        const endDate = c.req.query('endDate');

        const shopId = await getShopIdBySlug(pool, shopSlug);
        if (!shopId) {
            return c.json({ success: false, message: "Warung tidak ditemukan." }, 404);
        }

        const { results: shopRows } = await pool.prepare(
            `SELECT package_id FROM shops WHERE id = ?`
        ).bind(shopId).all();

        if (!shopRows || shopRows.length === 0) {
            return c.json({ success: false, message: "Data toko tidak ditemukan." }, 404);
        }

        const packageId = Number(shopRows[0].package_id);
        if (packageId !== 2 && packageId !== 3) {
            return c.json({
                success: false,
                message: "Fitur Laporan Laba Rugi hanya tersedia untuk Paket Juragan dan Sultan."
            }, 403);
        }

        let dateFilter = "AND DATE(o.created_at) = DATE('now')";
        let params = [shopId];

        if (startDate && endDate) {
            dateFilter = "AND DATE(o.created_at) BETWEEN ? AND ?";
            params.push(startDate, endDate);
        }

        const categoriesQuery = `
            SELECT 
                COALESCE(c.name, 'Tanpa Kategori') AS category_name,
                SUM(od.subtotal) AS total_gross_sales,
                SUM(od.qty * COALESCE(od.cost_price, p.cost_price, 0)) AS total_hpp,
                SUM(od.qty) AS total_qty_sold
            FROM order_details od
            JOIN orders o ON od.order_id = o.id
            LEFT JOIN products p ON od.product_id = p.id
            LEFT JOIN categories c ON p.category = c.name AND c.shop_id = o.shop_id
            WHERE o.shop_id = ? AND o.status = 'completed' ${dateFilter}
            GROUP BY c.name
            ORDER BY total_gross_sales DESC
        `;

        const { results: categoryRows } = await pool.prepare(categoriesQuery).bind(...params).all();

        const summaryQuery = `
            SELECT 
                COALESCE(SUM(subtotal), 0) AS total_gross_sales,
                COALESCE(SUM(discount), 0) AS total_discount,
                COALESCE(SUM(tax), 0) AS total_tax,
                COALESCE(SUM(total), 0) AS total_net_sales
            FROM orders o
            WHERE o.shop_id = ? AND o.status = 'completed' ${dateFilter}
        `;

        const { results: summaryRows } = await pool.prepare(summaryQuery).bind(...params).all();
        const summary = summaryRows[0];

        let grandTotalHpp = 0;
        const formattedCategories = categoryRows.map(row => {
            const grossSales = parseFloat(row.total_gross_sales) || 0;
            const hpp = parseFloat(row.total_hpp) || 0;
            const grossProfit = grossSales - hpp;
            grandTotalHpp += hpp;

            return {
                category_name: row.category_name,
                qty_sold: parseInt(row.total_qty_sold) || 0,
                gross_sales: grossSales,
                hpp: hpp,
                gross_profit: grossProfit
            };
        });

        const grossSalesTotal = parseFloat(summary.total_gross_sales) || 0;
        const discountTotal = parseFloat(summary.total_discount) || 0;
        const netSalesTotal = grossSalesTotal - discountTotal;
        const netProfit = netSalesTotal - grandTotalHpp;

        return c.json({
            success: true,
            data: {
                period: {
                    start_date: startDate || new Date().toISOString().split('T')[0],
                    end_date: endDate || new Date().toISOString().split('T')[0]
                },
                summary: {
                    total_gross_sales: grossSalesTotal,
                    total_discount: discountTotal,
                    total_net_sales: netSalesTotal,
                    total_hpp: grandTotalHpp,
                    net_profit: netProfit,
                    total_tax: parseFloat(summary.total_tax) || 0
                },
                categories_breakdown: formattedCategories
            }
        });

    } catch (error) {
        console.error("Error laporan laba rugi:", error);
        return c.json({ success: false, message: "Gagal memuat laporan laba rugi: " + error.message }, 500);
    }
});

app.get('/api/reports/profit-loss/export-excel', verifikasiAksesWarung, async (c) => {
    try {
        const pool = getDbPool(c);
        const shopSlug = c.req.query('shop') || c.req.header('x-shop-slug');
        const startDate = c.req.query('startDate');
        const endDate = c.req.query('endDate');

        const shopId = await getShopIdBySlug(pool, shopSlug);
        if (!shopId) {
            return c.json({ success: false, message: "Warung tidak ditemukan." }, 404);
        }

        const { results: shopRows } = await pool.prepare(`SELECT package_id FROM shops WHERE id = ?`).bind(shopId).all();
        if (!shopRows || shopRows.length === 0 || (Number(shopRows[0].package_id) !== 2 && Number(shopRows[0].package_id) !== 3)) {
            return c.json({ success: false, message: "Akses ditolak. Fitur khusus Paket Juragan dan Sultan." }, 403);
        }

        let dateFilter = "AND DATE(o.created_at) = DATE('now')";
        let params = [shopId];

        if (startDate && endDate) {
            dateFilter = "AND DATE(o.created_at) BETWEEN ? AND ?";
            params.push(startDate, endDate);
        }

        const categoriesQuery = `
            SELECT 
                COALESCE(c.name, 'Tanpa Kategori') AS category_name,
                SUM(od.subtotal) AS total_gross_sales,
                SUM(od.qty * COALESCE(od.cost_price, p.cost_price, 0)) AS total_hpp,
                SUM(od.qty) AS total_qty_sold
            FROM order_details od
            JOIN orders o ON od.order_id = o.id
            LEFT JOIN products p ON od.product_id = p.id
            LEFT JOIN categories c ON p.category = c.name AND c.shop_id = o.shop_id
            WHERE o.shop_id = ? AND o.status = 'completed' ${dateFilter}
            GROUP BY c.name
            ORDER BY total_gross_sales DESC
        `;
        const { results: categoryRows } = await pool.prepare(categoriesQuery).bind(...params).all();

        const summaryQuery = `
            SELECT 
                COALESCE(SUM(subtotal), 0) AS total_gross_sales,
                COALESCE(SUM(discount), 0) AS total_discount,
                COALESCE(SUM(tax), 0) AS total_tax,
                COALESCE(SUM(total), 0) AS total_net_sales
            FROM orders o
            WHERE o.shop_id = ? AND o.status = 'completed' ${dateFilter}
        `;
        const { results: summaryRows } = await pool.prepare(summaryQuery).bind(...params).all();
        const summary = summaryRows[0];

        let grandTotalHpp = 0;
        categoryRows.forEach(row => {
            grandTotalHpp += parseFloat(row.total_hpp) || 0;
        });

        const grossSalesTotal = parseFloat(summary.total_gross_sales) || 0;
        const discountTotal = parseFloat(summary.total_discount) || 0;
        const netSalesTotal = grossSalesTotal - discountTotal;
        const netProfit = netSalesTotal - grandTotalHpp;

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Laporan Laba Rugi');

        worksheet.mergeCells('A1:E1');
        worksheet.getCell('A1').value = 'LAPORAN LABA RUGI';
        worksheet.getCell('A1').font = { size: 14, bold: true };

        worksheet.mergeCells('A2:E2');
        worksheet.getCell('A2').value = `Periode: ${startDate || 'Hari Ini'} s/d ${endDate || 'Hari Ini'}`;
        worksheet.getCell('A2').font = { size: 10, italic: true };

        worksheet.addRow([]);

        worksheet.addRow(['RINGKASAN FINANSIAL', 'NILAI (RP)']);
        worksheet.getRow(4).font = { bold: true, color: { argb: 'FFFFFF' } };
        worksheet.getRow(4).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D97706' } };

        worksheet.addRow(['Penjualan Kotor', grossSalesTotal]);
        worksheet.addRow(['Total Diskon Penjualan', discountTotal]);
        worksheet.addRow(['Penjualan Bersih', netSalesTotal]);
        worksheet.addRow(['Total HPP (Harga Pokok Penjualan)', grandTotalHpp]);
        worksheet.addRow(['LABA BERSIH OPERASIONAL', netProfit]);

        worksheet.getRow(9).font = { bold: true };

        worksheet.addRow([]);

        worksheet.addRow(['RINCIAN PER KATEGORI MASTER']);
        worksheet.getRow(11).font = { bold: true, size: 11 };

        const catHeaderRow = worksheet.addRow(['Kategori', 'Terjual (Qty)', 'Penjualan Kotor (Rp)', 'Total HPP (Rp)', 'Laba Kotor (Rp)']);
        catHeaderRow.font = { bold: true, color: { argb: 'FFFFFF' } };
        catHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '475569' } };

        categoryRows.forEach(c => {
            const gross = parseFloat(c.total_gross_sales) || 0;
            const hpp = parseFloat(c.total_hpp) || 0;
            worksheet.addRow([
                c.category_name,
                parseInt(c.total_qty_sold) || 0,
                gross,
                hpp,
                gross - hpp
            ]);
        });

        worksheet.columns = [
            { width: 28 },
            { width: 15 },
            { width: 22 },
            { width: 20 },
            { width: 20 }
        ];

        const buffer = await workbook.xlsx.writeBuffer();
        return c.body(buffer, 200, {
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Content-Disposition': `attachment; filename="Laporan_Laba_Rugi_${shopSlug}_${Date.now()}.xlsx"`
        });
    } catch (error) {
        console.error("Error export excel laba rugi:", error);
        return c.json({ success: false, message: "Gagal mengeksport laporan laba rugi." }, 500);
    }
});

// ---------------- GEMINI AI ASSISTANT ----------------
app.post('/api/tanya-ai', async (c) => {
    try {
        const { message, history } = await c.req.json();

        if (!message) {
            return c.json({ success: false, message: "Pesan tidak boleh kosong." }, 400);
        }

        const apiKey = c.env.GEMINI_API_KEY;
        if (!apiKey) {
            return c.json({ success: false, message: "API Key Gemini belum diatur di Variables Cloudflare/Environment." }, 500);
        }

        const ai = new GoogleGenAI({ apiKey: apiKey });

        const systemInstruction = `
        Kamu adalah Asisten BEDA. Tugasmu membantu pemilik warung/usaha terkait settingan warung, QRIS, kelola stok, dan proses transaksi. Jawab dengan bahasa ramah, sopan, dan solutif. Jika ada kendala teknis darurat, sarankan hubungi WhatsApp Admin BEDApos di 089525147422 atau email support@bedadigital.app .
        ATURAN GAYA MENJAWAB (SANGAT PENTING):
        1. Jawab secara RINGKAS, PADAT, dan LANGSUNG KE INTI (Maksimal 2 - 3 kalimat).
        2. DILARANG menggunakan basa-basi pembuka yang panjang (seperti "Halo! Senang sekali bisa membantu Anda...").
        3. Jangan gunakan poin-poin panjang kecuali diminta secara spesifik oleh pengguna.
        4. Gunakan bahasa Indonesia sehari-hari yang ramah, sopan, dan mudah dipahami pemilik warung.
        Informasi Penting :
        - untuk pendafaran klik tombol Daftar Sekarang di web bedapos.bedadigital.app
        - Ada live demo nya,  klik tombol Live Demo POS (login nya no.HP: 012345678 password: demo123)
        - Pendaftaran gratis trial 14 hari via bedapos.bedadigital.app.
        - Uang hasil penjualan 100% masuk ke rekening/QRIS pribadi pemilik warung (0% komisi).
        - Untuk tata cara penggunakan bisa dilihat di panduan
        Informasi tambahan :
        - BEDApos cukup diakses via browser HP/Laptop tanpa perlu install aplikasi.
        - Uang pembayaran dari pelanggan akan langsung masuk ke rekening bank atau QRIS pribadi milik Anda sendiri tanpa melalui pihak ketiga.
        - kalau butuh panduan tata cara, klik tombol PANDUAN di web ini
        
        Apa yang harus dilakukan setelah daftar.
        1.	Masuk ke Alamat web https://pos.bedadigital.app/login.html , lakukan login
        2.	Masuk ke Pengaturan, kalau dari HP,klik tombol titik 3, klik pengaturan
        3.	Setting yang diperlukan
        -	Ada Kena Pajak?, ceklis kalau memang ada pajak, ketik nilai persen pajaknya, kasih 0 jika barang/jasa mu sudah termasuk pajak. Kalau usaha mu belum ada pajak maka biarkan tidak ter ceklis.
        -	Hitung Stok Otomatis?, ini berlaku untuk paket Juragan dan Sultan, ceklis,jika transaksi barang mu ada pengecekan stok, sehingga stok yang sudah 0 tidak bisa di transaksi. Ceklis nya hilangkan jika memang belum siap untuk menerapkan hitung stok otomatis. Untuk paket UMKM hitung stok otomatis tidak ada, jadi murni transaksi tanpa melihat stok.
        -	Diskon standard toko (%), ini diisi apabila anda memberikan diskon di setiap transaksi yang terjadi, misalkan pada waktu-waktu tertentu, maka apabila ini diisi, setiap transaksi yang terjadi akan terpotong discount ini. Jika sudah tidak diperlukan lagi diskon ini, maka isi dengan angka 0.
        -	Rincian Akun bank, Isi no. rekening usaha anda disini, supaya nanti di kasir bisa langsung dilihat no. rekening nya apabila ada yang menggunakan metode pembayaran transfer.
        -	Ganti foto QRIS, jika memiliki QIRS, upload gambar QRIS mu disitu, sehingga di kasir bisa langsung tampil dan bisa langsung di scan.
        -	Klik simpan kalau sudah selesai.
        4.	Kelola Produk, Jika menggunakan HP, didashboard, klik tombol titik 3,klik Kelola Produk
        -	Export Excel, untuk export ke excel daftar produk dan stok terakhir usaha anda.
        -	Kategori , setting kategori-kategori produk usaha anda, misalkan kategori BARANG, JASA, atau lebih spesifik lagi, MAKANAN, MINUMAN, SPAREPART, JASA.
        -	Tambah Produk Baru, untuk menambah produk baru

        Untuk barcode itu opsional,bisa langsung discan dari HP nya, dengan klik tombol gambar kamera. Harga Modal/HPP isi untuk menentukan rugi laba. Stok isi apabila menggunakan hitung stok otomatis.
        Proses Transaksi di aplikasi POS
        1.	Transaksi POS,klik tombol POS di dashboar.
        2.	list katalog produk nya, bisa scan barcode pake HP dengan klik tombol gambar kamera di pojok pencarian barang.
        3.	Ini Keranjang untuk proses pembayaran,pembayaran bisa Tunai, QRIS dan Transfer. Untuk QRIS dan Transfer harus di upload atau di foto bukti bayarnya, langsung dari HP untuk foto bukti bayar nya. Ketika klik proses Order, maka akan muncul struk. 
        4.	Pembayaran QRIS, upload/foto bukti bayar
        5.	Pembayaran Transfer,upload / foto bukti bayar
       

        Cara Sambungkan ke Printer Bluetooth
        Panduan menghubungkan aplikasi ke printer thermal Bluetooth
        1. Aktifkan Bluetooth & Hubungkan Printer ke HP
        Pastikan printer sudah menyala dan Bluetooth di HP Anda sudah aktif. Scan/cari perangkat printer hingga muncul dan terhubung ke HP Anda.
        Passcode umum: 0000
        2.Instal Aplikasi RawBT
        Instal aplikasi RawBT inkless print service (gratis) melalui Google Play Store.
        3. Buka Pengaturan RawBT
        Buka aplikasi RawBT, lalu klik ikon Setting (Pengaturan).
        4. Tambah Printer (ADD PRINTER)
        Pada menu Settings, klik tombol ADD PRINTER
        5. Pilih Metode Bluetooth
        Pilih metode koneksi Bluetooth
        Pilih Koneksi Bluetooth
        6. Pilih Perangkat Printer
        Klik opsi not selected Pilih nama printer Bluetooth Anda (misal: RPP02N). Jika belum muncul, klik tombol SCANNING DEVICE
        7. Hubungkan Printer (CONNECT)
        Setelah perangkat printer dipilih, klik tombol CONNECT untuk menyelesaikan penyambungan.
    
        Paket
        UMKM
        Rp 25.000 / bulan
        atau Rp 225.000 / tahun
        - Maksimal 600 Transaksi/Bulan
        - Unlimited Produk
        - Cetak Struk Kasir
        - Laporan Penjualan + Export to Excel

        JURAGAN
        Rp 99.000 / bulan
        atau Rp 999.000 / tahun
        - Maksimal 3.000 Transaksi/Bulan
        - Unlimited Produk
        - Cetak Struk Kasir
        - Manajemen Stok
        - Hitung otomatis Harga HPP
        - Laporan Penjualan + Export Excel
        - Laporan Laba Rugi

        SULTAN
        Rp 211.000 / bulan
        atau Rp 2.110.000 / tahun
        - Transaksi Tanpa Batas (Unlimited)
        - Unlimited Produk
        - Cetak Struk Kasir
        - Manajemen Stok
        - Hitung otomatis Harga HPP
        - Laporan Penjualan + Export Excel
        - Laporan Laba Rugi
        `;

        const contentsPayload = [];

        if (history && Array.isArray(history)) {
            history.forEach(item => {
                contentsPayload.push({
                    role: item.role === 'user' ? 'user' : 'model',
                    parts: [{ text: item.text }]
                });
            });
        }

        contentsPayload.push({
            role: 'user',
            parts: [{ text: message }]
        });

        const response = await ai.models.generateContent({
            model: 'gemini-3.1-flash-lite',
            contents: contentsPayload,
            config: {
                systemInstruction: systemInstruction
            }
        });

        return c.json({
            success: true,
            reply: response.text
        });

    } catch (error) {
        console.error("Error Internal Tanya AI:", error);
        return c.json({ 
            success: false, 
            message: "Gagal memproses pesan AI: " + error.message 
        }, 500);
    }
});

// Serve file statis HTML/CSS/JS frontend dari folder
app.get('/*', serveStatic({ root: './' }));

export default app;