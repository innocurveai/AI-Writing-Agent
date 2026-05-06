-- =============================================================================
-- creative_writer_ai.py 용 Supabase 스키마
-- 적용: Supabase 대시보드 → SQL Editor → 아래 전체 실행
-- .env: SUPABASE_URL, SUPABASE_ANON_KEY (이 스크립트에는 넣지 않음)
-- =============================================================================

-- pgvector (이미 있으면 스킵)
create extension if not exists vector;

-- -----------------------------------------------------------------------------
-- 테이블
-- -----------------------------------------------------------------------------

create table if not exists public.user_documents (
  id uuid primary key default gen_random_uuid(),
  filename text not null,
  file_type text not null,
  upload_date timestamptz not null default now(),
  genre text null
);

comment on table public.user_documents is '업로드 원본 메타 (시/에세이/소설 태그는 genre)';
comment on column public.user_documents.genre is 'poetry | essay | novel 또는 null(공통)';

create table if not exists public.document_embeddings (
  id uuid primary key default gen_random_uuid(),
  doc_id uuid not null references public.user_documents (id) on delete cascade,
  chunk_content text not null,
  embedding_vector vector(1536) not null
);

comment on table public.document_embeddings is '청크 텍스트 + OpenAI text-embedding-3-small 벡터';
comment on column public.document_embeddings.embedding_vector is '1536차원, 코사인 거리(<=>) 검색';

-- -----------------------------------------------------------------------------
-- RLS (anon 키로 Streamlit에서 접근할 때 필요)
-- -----------------------------------------------------------------------------

alter table public.user_documents enable row level security;
alter table public.document_embeddings enable row level security;

drop policy if exists "anon_authenticated_rw_user_documents" on public.user_documents;
create policy "anon_authenticated_rw_user_documents"
  on public.user_documents
  for all
  to anon, authenticated
  using (true)
  with check (true);

drop policy if exists "anon_authenticated_rw_document_embeddings" on public.document_embeddings;
create policy "anon_authenticated_rw_document_embeddings"
  on public.document_embeddings
  for all
  to anon, authenticated
  using (true)
  with check (true);

-- -----------------------------------------------------------------------------
-- 유사도 검색 RPC (앱에서 supabase.rpc('match_document_embeddings', ...) 호출)
-- -----------------------------------------------------------------------------

create or replace function public.match_document_embeddings(
  query_embedding vector(1536),
  match_count int default 6,
  match_threshold float default 0.25,
  filter_doc_ids uuid[] default null,
  filter_genre text default null
)
returns table (
  id uuid,
  doc_id uuid,
  chunk_content text,
  similarity float
)
language sql
stable
security invoker
set search_path = public
as $$
  select
    de.id,
    de.doc_id,
    de.chunk_content,
    (1 - (de.embedding_vector <=> query_embedding))::float as similarity
  from public.document_embeddings de
  join public.user_documents ud on ud.id = de.doc_id
  where
    (filter_doc_ids is null or de.doc_id = any(filter_doc_ids))
    and (filter_genre is null or ud.genre is null or ud.genre = filter_genre)
    and (1 - (de.embedding_vector <=> query_embedding)) >= match_threshold
  order by de.embedding_vector <=> query_embedding
  limit greatest(match_count, 1);
$$;

-- -----------------------------------------------------------------------------
-- 권한 (API anon / 로그인 사용자가 테이블·RPC 사용)
-- -----------------------------------------------------------------------------

grant usage on schema public to anon, authenticated;

grant select, insert, update, delete on table public.user_documents to anon, authenticated;
grant select, insert, update, delete on table public.document_embeddings to anon, authenticated;

grant execute on function public.match_document_embeddings(
  vector,
  int,
  double precision,
  uuid[],
  text
) to anon, authenticated;

-- -----------------------------------------------------------------------------
-- (선택) 벡터 인덱스 — 데이터가 많아지면 주석 해제 후 실행 권장
-- -----------------------------------------------------------------------------

-- create index if not exists document_embeddings_embedding_hnsw_idx
-- on public.document_embeddings
-- using hnsw (embedding_vector vector_cosine_ops);

-- create index if not exists document_embeddings_embedding_ivfflat_idx
-- on public.document_embeddings
-- using ivfflat (embedding_vector vector_cosine_ops) with (lists = 100);
