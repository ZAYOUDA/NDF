-- 1. Créer la table des dépenses
create table expenses (
  id           serial primary key,
  description  text    default '',
  client       text    default '',
  projet       text    default '',
  transport    numeric default 0,
  repas        numeric default 0,
  commentaire  text    default '',
  vat_lines    jsonb   default '[]',
  total_ttc    numeric default 0,
  total_ht     numeric default 0,
  tva_2_6      numeric default 0,
  tva_5_5      numeric default 0,
  tva_10       numeric default 0,
  tva_20       numeric default 0,
  images       text[]  default '{}',
  created_at   timestamptz default now()
);

-- 2. Activer la sécurité RLS (désactivée pour service_role)
alter table expenses enable row level security;

-- 3. Créer le bucket de stockage pour les justificatifs
insert into storage.buckets (id, name, public)
values ('receipts', 'receipts', true);

-- 4. Politique : lecture publique des images
create policy "Public read receipts"
  on storage.objects for select
  using ( bucket_id = 'receipts' );

-- 5. Politique : upload via service_role uniquement
create policy "Service upload receipts"
  on storage.objects for insert
  with check ( bucket_id = 'receipts' );

create policy "Service delete receipts"
  on storage.objects for delete
  using ( bucket_id = 'receipts' );
