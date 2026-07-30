-- V2 更新脚本
-- 先在 Supabase SQL Editor 中完整运行，再部署新版 app.py。

-- 1) 原耗材使用表增加货主名称。
alter table public.consumable_usage
    add column if not exists owner_name text;

-- 2) 标准耗材主数据。
create table if not exists public.consumable_master (
    material_code text primary key,
    material_name text not null,
    material_size text,
    pieces_per_pallet numeric not null check (pieces_per_pallet > 0),
    active boolean not null default true,
    uploader text,
    updated_at timestamp without time zone default now()
);

-- 3) 每周库存记录。
create table if not exists public.consumable_inventory (
    inventory_key text primary key,
    inventory_date date not null,
    material_code text not null,
    material_name text not null,
    material_size text,
    pieces_per_pallet numeric not null,
    pallet_qty numeric not null check (pallet_qty >= 0),
    inventory_pieces bigint not null check (inventory_pieces >= 0),
    recorder text not null,
    notes text,
    recorded_at timestamp without time zone default now(),
    updated_at timestamp without time zone default now()
);

create index if not exists idx_consumable_inventory_date
    on public.consumable_inventory (inventory_date);

create index if not exists idx_consumable_inventory_material
    on public.consumable_inventory (material_code);

alter table public.consumable_master enable row level security;
alter table public.consumable_inventory enable row level security;

-- Streamlit 使用服务器端 service_role / secret key。
grant usage on schema public to service_role;

grant select, insert, update, delete
on table public.consumable_usage
to service_role;

grant select, insert, update, delete
on table public.consumable_master
to service_role;

grant select, insert, update, delete
on table public.consumable_inventory
to service_role;
