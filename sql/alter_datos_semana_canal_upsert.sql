-- Migración: columna canal + UPSERT desde mapa.html / Netlify/supabase.js
-- Ejecutar UNA VEZ si ya creaste datos_semana con la versión antigua (sin canal).
-- Proyectos nuevos: basta con sql/create_tables.sql actualizado.

alter table public.datos_semana add column if not exists canal text;

update public.datos_semana
set canal = 'legacy_' || id::text
where canal is null or btrim(canal) = '';

create unique index if not exists datos_semana_canal_key on public.datos_semana (canal);

-- RLS: permitir publicar / actualizar la fila oficial con la clave anon (mapa + Supabase)
drop policy if exists datos_semana_insert_anon on public.datos_semana;
create policy datos_semana_insert_anon
  on public.datos_semana for insert to anon with check (true);

drop policy if exists datos_semana_update_anon on public.datos_semana;
create policy datos_semana_update_anon
  on public.datos_semana for update to anon using (true) with check (true);

-- Forzar recarga del caché de esquema de PostgREST (evita error "Could not find the 'canal' column ... schema cache")
notify pgrst, 'reload schema';
