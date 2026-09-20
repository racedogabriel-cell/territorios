-- Permitir que la app Netlify borre acciones aún no procesadas (cola compartida)
drop policy if exists acciones_delete_anon on public.acciones_colaboradores;
create policy acciones_delete_anon
  on public.acciones_colaboradores
  for delete
  to anon
  using (procesado = false);

notify pgrst, 'reload schema';
