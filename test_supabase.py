from supabase_service import SupabaseService

supa = SupabaseService()

proveedor = supa.buscar_proveedor("E45351186")

print(proveedor)