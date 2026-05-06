# 고유 문체 기반 AI Writing Agent

Streamlit으로 동작하는 한국어 창작 보조 앱입니다. **Supabase(pgvector)**에 올린 학습 문서의 임베딩과 **OpenAI**를 사용해, 키워드에 맞는 **시·에세이·소설** 초안을 **사용자 문체**에 가깝게 생성합니다.

## 필요한 것

- Python 3.10 이상 권장
- [Supabase](https://supabase.com/) 프로젝트(pgvector)
- [OpenAI](https://platform.openai.com/) API 키

## 빠른 시작 (Windows PowerShell)

### 1. 저장소 폴더로 이동

```powershell
cd <이 프로젝트가 있는 폴더 경로>
```

### 2. 가상환경과 패키지

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### 3. Supabase 스키마 적용

Supabase 대시보드 → **SQL Editor**에서 `supabase_schema.sql` 전체를 실행합니다.  
테이블(`user_documents`, `document_embeddings`), RLS 정책, RPC 함수 `match_document_embeddings`가 생성됩니다.

### 4. 환경 변수

프로젝트 루트에 `.env` 파일을 만들고 다음을 채웁니다. (이 파일은 `.gitignore`에 포함되어 Git에 올리지 마세요.)

| 변수 | 설명 |
|------|------|
| `SUPABASE_URL` | Supabase 프로젝트 URL |
| `SUPABASE_ANON_KEY` | Supabase anon(public) 키 |
| `OPENAI_API_KEY` | OpenAI API 키 |
| `OPENAI_CHAT_MODEL` | (선택) 채팅 모델 ID. 없으면 코드 기본값 사용 |

예시 형태만 참고하세요. 실제 키는 본인 것으로 교체합니다.

```env
SUPABASE_URL=https://xxxx.supabase.co
SUPABASE_ANON_KEY=eyJ...
OPENAI_API_KEY=sk-...
# OPENAI_CHAT_MODEL=gpt-4o-mini
```

### 5. 앱 실행

```powershell
streamlit run creative_writer_ai.py
```

브라우저에서 안내되는 주소(기본 `http://localhost:8501`)로 접속합니다.

## 사용 흐름

1. **사이드바**에서 학습용 문서(PDF, TXT, MD 등)를 업로드하고 장르를 지정합니다. 임베딩은 `text-embedding-3-small`(1536차원)을 사용합니다.
2. 중앙에서 **장르**, **키워드**, **문체 믹싱**(전체 / 선택 작품만), **문체 생성 옵션**(유사도 0~1)을 설정합니다.
3. **초안 생성**을 누르면 벡터 검색으로 스타일 레퍼런스를 모은 뒤, OpenAI 채팅 API로 초안을 작성합니다.

유사도는 **1에 가까울수록** 학습 문서와 임베딩이 비슷한 청크 위주로, **0에 가까울수록** 참고 범위가 넓어질 수 있습니다.

## LangChain은 어디에 쓰이나요?

이 프로젝트는 LangChain의 **체인(Chain), 에이전트, LCEL, Retriever** 같은 상위 프레임워크는 **사용하지 않습니다.**  
임베딩·채팅·Supabase RPC 호출은 **`openai`**, **`supabase`** Python 클라이언트로 직접 처리합니다.

다만 **학습 문서를 잘게 나누는(청킹) 단계**만 LangChain 생태계의 **`langchain_text_splitters`** 패키지를 씁니다.

| 구분 | 사용 여부 |
|------|-----------|
| `langchain_text_splitters.RecursiveCharacterTextSplitter` | **사용** — 업로드 텍스트를 청크로 분할 |
| `langchain` 본체, `langchain_openai` 등 | 코드에서 **직접 import 하지 않음** (`requirements.txt`에 포함될 수 있으나, 앱 로직의 핵심 경로에는 없음) |

즉, “LangChain을 쓴다”고 하면 **텍스트 스플리터 유틸 한 종류** 정도로 이해하면 됩니다.

## 텍스트 청킹은 어떻게 동작하나요?

### 언제 청킹이 일어나나요?

**사이드바에서 학습 파일을 업로드하고 처리할 때**입니다. PDF·TXT·MD·JSON에서 본문을 추출한 뒤, 그 문자열을 청크 목록으로 나눕니다.

### 어떤 규칙으로 나뉘나요?

`creative_writer_ai.py`의 `chunk_text()`에서 **`RecursiveCharacterTextSplitter`**를 사용합니다.

- **`chunk_size=900`** — 한 청크당 대략 900자(문자 개수 기준)를 넘기지 않게 잘라, 임베딩 입력 길이와 문맥 단위를 맞춥니다.
- **`chunk_overlap=140`** — 인접 청크끼리 140자 정도 겹칩니다. 문장이 청크 경계에서 끊기면 문맥이 잘리기 쉬운데, 겹침으로 그 손실을 줄입니다.
- **`length_function=len`** — 길이는 Python 문자열의 글자 수로 셉니다.

나눈 뒤 빈 문자열은 버리고, 각 청크마다 **OpenAI `text-embedding-3-small`**으로 벡터를 만든 다음 **`document_embeddings`** 테이블에 `chunk_content`와 함께 저장합니다.

### 청킹이 검색·생성에 어떻게 연결되나요?

- DB와 벡터 검색의 단위는 **“원본 파일 전체”가 아니라 “청크 한 덩어리”**입니다.
- 사용자가 키워드를 넣고 초안을 만들면, RPC **`match_document_embeddings`**가 질문 임베딩과 비슷한 **청크들**을 고릅니다.
- 그 청크 텍스트들이 프롬프트의 **스타일 레퍼런스**로 들어가고, 채팅 API가 그 문체를 참고해 초안을 씁니다.

검색 시 **몇 개의 청크**를 가져올지는 코드 상수 `STYLE_REFERENCE_TOP_K`(기본 6)로 정합니다. 청크 **크기·겹침**을 바꾸려면 같은 파일의 `chunk_text()` 안 `RecursiveCharacterTextSplitter` 인자를 수정하면 됩니다.

## 프로젝트 구조

| 경로 | 설명 |
|------|------|
| `creative_writer_ai.py` | Streamlit UI, 업로드·청킹·임베딩·검색·LLM 호출 |
| `supabase_schema.sql` | DB 테이블, RLS, `match_document_embeddings` RPC |
| `requirements.txt` | Python 의존성 |
| `.env` | 비밀 키(로컬 전용, 저장소에 커밋하지 않음) |

검색 시 가져올 레퍼런스 청크 개수(`STYLE_REFERENCE_TOP_K`)는 위 **«텍스트 청킹은 어떻게 동작하나요?»** 절을 참고하세요.

## 문제 해결

- **벡터 검색 오류**: `supabase_schema.sql`이 프로젝트에 적용됐는지, 컬럼명이 `embedding_vector`인지 확인하세요.
- **빈 초안 / 모델 오류**: `.env`에 `OPENAI_CHAT_MODEL`을 계정에서 사용 가능한 모델 ID로 지정해 보세요.
- **스타일 레퍼런스 없음**: 학습 문서 업로드·처리 후 임베딩이 들어갔는지, 유사도를 낮추거나 문서 선택을 조정해 보세요.

## 라이선스

이 저장소에 별도 라이선스 파일이 없다면, 사용·배포 조건은 저장소 소유자에게 문의하세요.
