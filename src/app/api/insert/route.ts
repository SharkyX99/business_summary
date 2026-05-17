import { NextResponse } from "next/server";
import { createClient } from "@supabase/supabase-js";

// ฟังก์ชันสำหรับสร้าง Supabase Client เมื่อมีการเรียกใช้งาน API เท่านั้น
function getSupabaseClient() {
  const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;
  const supabaseAnonKey = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;

  if (!supabaseUrl || !supabaseAnonKey) {
    throw new Error(
      "❌ ไม่พบตั้งค่าเซิร์ฟเวอร์ฐานข้อมูล กรุณาตรวจสอบไฟล์ .env.local",
    );
  }

  return createClient(supabaseUrl, supabaseAnonKey);
}

// 1. API สำหรับดึงข้อมูลทั้งหมดมาแสดงและเปรียบเทียบบนหน้าจอ
export async function GET() {
  try {
    const supabase = getSupabaseClient();

    const { data, error } = await supabase
      .from("sales_data")
      .select("*")
      .order("sales_date", { ascending: false });

    if (error) throw error;
    return NextResponse.json(data);
  } catch (error: any) {
    return NextResponse.json(
      { success: false, error: error.message },
      { status: 500 },
    );
  }
}

// 2. API สำหรับรับข้อมูลจากไฟล์ Excel แล้วบันทึกลงฐานข้อมูลแบบถาวร
export async function POST(request: Request) {
  try {
    const supabase = getSupabaseClient();
    const body = await request.json();

    const { data, error } = await supabase.from("sales_data").insert(body);

    if (error) throw error;
    return NextResponse.json({ success: true });
  } catch (error: any) {
    return NextResponse.json(
      { success: false, error: error.message },
      { status: 500 },
    );
  }
}
// 3. API สำหรับลบข้อมูลออกจากฐานข้อมูล
export async function DELETE(req: Request) {
  try {
    const supabase = getSupabaseClient();

    const body = await req.json();
    const { sales_date } = body;

    if (!sales_date) {
      return NextResponse.json(
        {
          success: false,
          error: "ไม่พบ sales_date",
        },
        { status: 400 },
      );
    }

    const { error } = await supabase
      .from("sales_data")
      .delete()
      .eq("sales_date", sales_date);

    if (error) throw error;

    return NextResponse.json({
      success: true,
    });
  } catch (error: any) {
    return NextResponse.json(
      {
        success: false,
        error: error.message,
      },
      { status: 500 },
    );
  }
}