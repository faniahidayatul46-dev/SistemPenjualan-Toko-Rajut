-- Little Stitches | Supabase PostgreSQL
-- ERD tepat 4 entitas: customers, products, sales, sale_items.

create extension if not exists pgcrypto;

create table if not exists public.customers (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  phone text,
  email text,
  address text,
  created_at timestamptz not null default now()
);

create table if not exists public.products (
  id uuid primary key default gen_random_uuid(),
  sku text not null unique,
  name text not null,
  category text not null default 'Aksesori Rajut',
  description text,
  price numeric(14,2) not null default 0 check (price >= 0),
  stock integer not null default 0 check (stock >= 0),
  minimum_stock integer not null default 3 check (minimum_stock >= 0),
  image_url text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.sales (
  id uuid primary key default gen_random_uuid(),
  invoice_no text not null unique,
  customer_id uuid references public.customers(id) on delete restrict,
  entry_type text not null default 'sale' check (entry_type in ('sale', 'purchase', 'stock_adjustment', 'settlement')),
  journal_description text,
  related_sale_id uuid references public.sales(id) on delete cascade,
  stock_product_id uuid references public.products(id) on delete restrict,
  stock_movement text check (stock_movement in ('in', 'damaged', 'lost')),
  stock_quantity integer check (stock_quantity > 0),
  stock_before integer check (stock_before >= 0),
  stock_after integer check (stock_after >= 0),
  sale_date date not null default current_date,
  sales_channel text not null default 'whatsapp_online' check (sales_channel in ('in_store', 'whatsapp_online')),
  delivery_method text not null default 'delivery' check (delivery_method in ('delivery', 'pickup')),
  payment_method text check (payment_method in ('transfer', 'qris')),
  payment_reference text,
  payment_proof_url text,
  payment_status text not null default 'pending' check (payment_status in ('pending', 'verified', 'rejected')),
  subtotal numeric(14,2) not null default 0 check (subtotal >= 0),
  shipping_fee numeric(14,2) not null default 15000 check (shipping_fee >= 0),
  total numeric(14,2) not null default 0 check (total >= 0),
  paid_amount numeric(14,2) not null default 0 check (paid_amount >= 0),
  notes text,
  created_at timestamptz not null default now()
);

alter table public.sales add column if not exists payment_reference text;
alter table public.sales add column if not exists paid_amount numeric(14,2) not null default 0;
alter table public.sales add column if not exists entry_type text not null default 'sale';
alter table public.sales add column if not exists journal_description text;
alter table public.sales add column if not exists related_sale_id uuid references public.sales(id) on delete cascade;
alter table public.sales add column if not exists delivery_method text not null default 'delivery';
alter table public.sales add column if not exists subtotal numeric(14,2) not null default 0;
alter table public.sales add column if not exists shipping_fee numeric(14,2) not null default 15000;
alter table public.sales add column if not exists stock_product_id uuid references public.products(id) on delete restrict;
alter table public.sales add column if not exists stock_movement text;
alter table public.sales add column if not exists stock_quantity integer;
alter table public.sales add column if not exists stock_before integer;
alter table public.sales add column if not exists stock_after integer;
alter table public.sales add column if not exists sales_channel text not null default 'whatsapp_online';
alter table public.sales alter column customer_id drop not null;
alter table public.sales alter column payment_method drop not null;
alter table public.sales drop constraint if exists sales_entry_type_check;
alter table public.sales add constraint sales_entry_type_check
  check (entry_type in ('sale', 'purchase', 'stock_adjustment', 'settlement'));
alter table public.sales drop constraint if exists sales_stock_movement_check;
alter table public.sales add constraint sales_stock_movement_check
  check (stock_movement is null or stock_movement in ('in', 'damaged', 'lost'));
alter table public.sales drop constraint if exists sales_stock_quantity_check;
alter table public.sales add constraint sales_stock_quantity_check
  check (stock_quantity is null or stock_quantity > 0);
alter table public.sales drop constraint if exists sales_stock_before_check;
alter table public.sales add constraint sales_stock_before_check
  check (stock_before is null or stock_before >= 0);
alter table public.sales drop constraint if exists sales_stock_after_check;
alter table public.sales add constraint sales_stock_after_check
  check (stock_after is null or stock_after >= 0);
alter table public.sales drop constraint if exists sales_sales_channel_check;
alter table public.sales add constraint sales_sales_channel_check
  check (sales_channel in ('in_store', 'whatsapp_online'));

create table if not exists public.sale_items (
  id uuid primary key default gen_random_uuid(),
  sale_id uuid not null references public.sales(id) on delete cascade,
  product_id uuid not null references public.products(id) on delete restrict,
  quantity integer not null check (quantity > 0),
  unit_price numeric(14,2) not null check (unit_price >= 0),
  subtotal numeric(14,2) generated always as (quantity * unit_price) stored,
  unique (sale_id, product_id)
);

create index if not exists products_stock_idx on public.products(stock);
create index if not exists sales_date_idx on public.sales(sale_date desc);
create index if not exists sale_items_sale_idx on public.sale_items(sale_id);
create index if not exists sales_related_sale_idx on public.sales(related_sale_id);

create or replace function public.record_stock_movement(
  p_product_id uuid,
  p_movement text,
  p_quantity integer,
  p_notes text default null
)
returns public.sales
language plpgsql
security invoker
set search_path = public
as $$
declare
  v_product public.products%rowtype;
  v_entry public.sales%rowtype;
  v_stock_before integer;
  v_stock_delta integer;
  v_movement_label text;
begin
  if p_quantity is null or p_quantity <= 0 then
    raise exception 'Jumlah unit harus lebih dari nol';
  end if;

  if p_movement is null or p_movement not in ('in', 'damaged', 'lost') then
    raise exception 'Jenis penyesuaian stok tidak valid';
  end if;

  select * into v_product
  from public.products
  where id = p_product_id
  for update;

  if not found then
    raise exception 'Produk tidak ditemukan';
  end if;

  v_stock_before := v_product.stock;
  v_stock_delta := case when p_movement = 'in' then p_quantity else -p_quantity end;
  if v_stock_before + v_stock_delta < 0 then
    raise exception 'Jumlah penyesuaian melebihi stok yang tersedia';
  end if;

  update public.products
  set stock = stock + v_stock_delta, updated_at = now()
  where id = p_product_id
  returning * into v_product;

  v_movement_label := case p_movement
    when 'in' then 'Stok masuk'
    when 'damaged' then 'Produk rusak'
    else 'Produk hilang'
  end;

  insert into public.sales (
    invoice_no, entry_type, journal_description, stock_product_id,
    stock_movement, stock_quantity, stock_before, stock_after,
    sale_date, delivery_method, payment_status, subtotal,
    shipping_fee, total, paid_amount, notes
  ) values (
    'STK-' || upper(substr(replace(gen_random_uuid()::text, '-', ''), 1, 12)),
    'stock_adjustment',
    format('%s · %s × %s unit', v_movement_label, v_product.name, p_quantity),
    p_product_id,
    p_movement,
    p_quantity,
    v_stock_before,
    v_product.stock,
    current_date,
    'pickup',
    'verified',
    0,
    0,
    0,
    0,
    nullif(trim(p_notes), '')
  )
  returning * into v_entry;

  return v_entry;
end;
$$;

create or replace function public.record_sale_settlement(
  p_sale_id uuid,
  p_amount numeric,
  p_payment_method text,
  p_payment_reference text default null,
  p_payment_proof_url text default null,
  p_payment_date date default current_date
)
returns public.sales
language plpgsql
security invoker
set search_path = public
as $$
declare
  v_sale public.sales%rowtype;
  v_settled numeric(14,2);
  v_outstanding numeric(14,2);
  v_entry public.sales%rowtype;
begin
  if p_amount is null or p_amount <= 0 then
    raise exception 'Nominal pelunasan harus lebih dari nol';
  end if;

  if p_payment_method is null or p_payment_method not in ('transfer', 'qris') then
    raise exception 'Metode pembayaran tidak valid';
  end if;

  if p_payment_date is null then
    raise exception 'Tanggal pembayaran wajib diisi';
  end if;

  select * into v_sale
  from public.sales
  where id = p_sale_id and entry_type = 'sale'
  for update;

  if not found then
    raise exception 'Transaksi penjualan tidak ditemukan';
  end if;

  select coalesce(sum(paid_amount), 0) into v_settled
  from public.sales
  where related_sale_id = p_sale_id
    and entry_type = 'settlement'
    and payment_status <> 'rejected';

  v_outstanding := v_sale.total
    - case when v_sale.payment_status = 'rejected' then 0 else v_sale.paid_amount end
    - v_settled;
  if p_amount > v_outstanding then
    raise exception 'Nominal pelunasan melebihi sisa tagihan sebesar %', v_outstanding;
  end if;

  insert into public.sales (
    invoice_no, customer_id, entry_type, journal_description, related_sale_id,
    sale_date, sales_channel, delivery_method, payment_method,
    payment_reference, payment_proof_url, payment_status,
    subtotal, shipping_fee, total, paid_amount, notes
  ) values (
    'SET-' || upper(substr(replace(gen_random_uuid()::text, '-', ''), 1, 12)),
    v_sale.customer_id,
    'settlement',
    'Pelunasan invoice ' || v_sale.invoice_no,
    v_sale.id,
    p_payment_date,
    v_sale.sales_channel,
    v_sale.delivery_method,
    p_payment_method,
    nullif(trim(p_payment_reference), ''),
    nullif(trim(p_payment_proof_url), ''),
    'pending',
    p_amount,
    0,
    p_amount,
    p_amount,
    null
  )
  returning * into v_entry;

  return v_entry;
end;
$$;

grant execute on function public.record_stock_movement(uuid, text, integer, text) to anon, authenticated;
grant execute on function public.record_sale_settlement(uuid, numeric, text, text, text, date) to anon, authenticated;

alter table public.customers enable row level security;
alter table public.products enable row level security;
alter table public.sales enable row level security;
alter table public.sale_items enable row level security;

drop policy if exists "public read customers" on public.customers;
drop policy if exists "public write customers" on public.customers;
drop policy if exists "public read products" on public.products;
drop policy if exists "public write products" on public.products;
drop policy if exists "public read sales" on public.sales;
drop policy if exists "public write sales" on public.sales;
drop policy if exists "public read sale items" on public.sale_items;
drop policy if exists "public write sale items" on public.sale_items;

create policy "public read customers" on public.customers for select using (true);
create policy "public write customers" on public.customers for all using (true) with check (true);
create policy "public read products" on public.products for select using (true);
create policy "public write products" on public.products for all using (true) with check (true);
create policy "public read sales" on public.sales for select using (true);
create policy "public write sales" on public.sales for all using (true) with check (true);
create policy "public read sale items" on public.sale_items for select using (true);
create policy "public write sale items" on public.sale_items for all using (true) with check (true);

-- Storage bucket untuk bukti pembayaran.
insert into storage.buckets (id, name, public)
values ('payment-proofs', 'payment-proofs', true)
on conflict (id) do update set public = true;

drop policy if exists "public read payment proofs" on storage.objects;
drop policy if exists "public upload payment proofs" on storage.objects;
drop policy if exists "public update payment proofs" on storage.objects;
drop policy if exists "public delete payment proofs" on storage.objects;

create policy "public read payment proofs" on storage.objects
  for select using (bucket_id = 'payment-proofs');
create policy "public upload payment proofs" on storage.objects
  for insert with check (bucket_id = 'payment-proofs');
create policy "public update payment proofs" on storage.objects
  for update using (bucket_id = 'payment-proofs') with check (bucket_id = 'payment-proofs');
create policy "public delete payment proofs" on storage.objects
  for delete using (bucket_id = 'payment-proofs');

insert into public.products (sku, name, category, description, price, stock, minimum_stock)
values
  ('LS-001', 'Tas Rajut Daisy', 'Tas Rajut', 'Tas tangan dengan detail bunga yang lembut.', 185000, 12, 4),
  ('LS-002', 'Bucket Hat Cozy', 'Fashion Rajut', 'Topi rajut hangat dengan siluet klasik.', 95000, 8, 3),
  ('LS-003', 'Cardigan Earth Tone', 'Busana Rajut', 'Cardigan nyaman untuk gaya sehari-hari.', 275000, 3, 5)
on conflict (sku) do nothing;

notify pgrst, 'reload schema';
