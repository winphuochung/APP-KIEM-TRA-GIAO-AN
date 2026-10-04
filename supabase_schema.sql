-- =========================================================================
-- HỆ THỐNG CƠ SỞ DỮ LIỆU SUPABASE CLOUD - HỆ SINH THÁI KIỂM ĐỊNH GIÁO DỤC SỐ
-- Trường TH & THCS Phước Hưng - Tổ Toán - KHTN - Công nghệ
-- =========================================================================

-- 1. Bảng lưu Sổ tay Tổ trưởng (Theo dõi giáo án & giảng dạy hàng tuần)
CREATE TABLE IF NOT EXISTS public.notebook_records (
    id BIGSERIAL PRIMARY KEY,
    week INT NOT NULL,
    teacher_id TEXT NOT NULL,
    teacher_name TEXT NOT NULL,
    so_tiet_ngay TEXT DEFAULT '',
    so_tiet_tuan INT DEFAULT 0,
    giao_an_drive TEXT DEFAULT 'Chưa nộp',
    dddh INT DEFAULT 0,
    cntt INT DEFAULT 0,
    nhan_xet TEXT DEFAULT '',
    xep_loai TEXT DEFAULT 'Tốt',
    updated_at TIMESTAMPTZ DEFAULT NOW(),
    CONSTRAINT uq_notebook_week_teacher UNIQUE (week, teacher_id)
);

-- 2. Bảng Danh sách Giáo viên Tổ Chuyên Môn
CREATE TABLE IF NOT EXISTS public.teachers (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    role TEXT DEFAULT 'Giáo viên',
    subject_str TEXT DEFAULT '',
    grades TEXT DEFAULT '',
    classes JSONB DEFAULT '[]'::jsonb,
    homeroom TEXT DEFAULT '',
    default_cntt INT DEFAULT 0,
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- 3. Bảng Nhật ký Giám sát Google Drive
CREATE TABLE IF NOT EXISTS public.drive_logs (
    id BIGSERIAL PRIMARY KEY,
    scanned_at TIMESTAMPTZ DEFAULT NOW(),
    total_teachers INT DEFAULT 8,
    teachers_data JSONB DEFAULT '[]'::jsonb
);

-- 4. Bảng Kết quả Kiểm tra Giáo án
CREATE TABLE IF NOT EXISTS public.inspection_records (
    id BIGSERIAL PRIMARY KEY,
    period TEXT NOT NULL,
    teacher_id TEXT NOT NULL,
    teacher_name TEXT NOT NULL,
    total_lessons INT DEFAULT 0,
    passed_lessons INT DEFAULT 0,
    lessons_data JSONB DEFAULT '[]'::jsonb,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

-- Cho phép truy cập RLS (Row Level Security) công khai cho Anon & Service Role
ALTER TABLE public.notebook_records ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.teachers ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.drive_logs ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.inspection_records ENABLE ROW LEVEL SECURITY;

-- DROP POLICY cũ nếu đã tồn tại để tránh lỗi 42710 (policy already exists)
DROP POLICY IF EXISTS "Allow public select on notebook_records" ON public.notebook_records;
DROP POLICY IF EXISTS "Allow public insert/update on notebook_records" ON public.notebook_records;

DROP POLICY IF EXISTS "Allow public select on teachers" ON public.teachers;
DROP POLICY IF EXISTS "Allow public insert/update on teachers" ON public.teachers;

DROP POLICY IF EXISTS "Allow public select on drive_logs" ON public.drive_logs;
DROP POLICY IF EXISTS "Allow public insert/update on drive_logs" ON public.drive_logs;

DROP POLICY IF EXISTS "Allow public select on inspection_records" ON public.inspection_records;
DROP POLICY IF EXISTS "Allow public insert/update on inspection_records" ON public.inspection_records;

-- TẠO POLICY MỚI CHUẨN ĐỊNH DẠNG
CREATE POLICY "Allow public select on notebook_records" ON public.notebook_records FOR SELECT USING (true);
CREATE POLICY "Allow public insert/update on notebook_records" ON public.notebook_records FOR ALL USING (true);

CREATE POLICY "Allow public select on teachers" ON public.teachers FOR SELECT USING (true);
CREATE POLICY "Allow public insert/update on teachers" ON public.teachers FOR ALL USING (true);

CREATE POLICY "Allow public select on drive_logs" ON public.drive_logs FOR SELECT USING (true);
CREATE POLICY "Allow public insert/update on drive_logs" ON public.drive_logs FOR ALL USING (true);

CREATE POLICY "Allow public select on inspection_records" ON public.inspection_records FOR SELECT USING (true);
CREATE POLICY "Allow public insert/update on inspection_records" ON public.inspection_records FOR ALL USING (true);
