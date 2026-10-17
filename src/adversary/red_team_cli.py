# -*- coding: utf-8 -*-
"""
Interactive Terminal Red Team - The Adversary CLI Console
Giao diện dòng lệnh độc lập cho Kẻ tấn công (The Adversary - Harry Halpin 2025).
Không dùng emoji hoặc icon, chuẩn học thuật, hiển thị màu sắc ANSI chuyên nghiệp.
"""
import os
import sys
import time

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../..")))

from src.adversary.adversary_engine import AdversaryEngine

# ANSI Colors
RED = "\033[91m"
GREEN = "\033[92m"
YELLOW = "\033[93m"
BLUE = "\033[94m"
MAGENTA = "\033[95m"
CYAN = "\033[96m"
BOLD = "\033[1m"
RESET = "\033[0m"


def print_banner():
    banner = f"""
{RED}{BOLD}===================================================================================
   RED TEAM: THE ADVERSARY ATTACK CONSOLE
   Dua tren bai bao: The Adversary - Harry Halpin (Oxford Journal of Cybersecurity 2025)
   He thong tan cong doi khang thuc chien doc lap voi Blue Team
==================================================================================={RESET}
    """
    print(banner)


def main():
    print_banner()
    engine = AdversaryEngine()

    print(f"[*] {CYAN}Muc tieu Blue Team:{RESET} {BOLD}{engine.target_dir}{RESET}")
    print(f"[*] {CYAN}Tong so tep phat hien:{RESET} {len(engine.get_target_files())} tep tin\n")

    while True:
        status = engine.get_status()
        print(f"{YELLOW}--- [TRANG THAI TAN CONG HIEN TAI] ---{RESET}")
        print(f"  - Tep bi sua doi (Tampered) : {RED}{status['tampered_count']}{RESET} tep")
        print(f"  - Tep la xam nhap (Injected): {MAGENTA}{status['injected_count']}{RESET} tep")
        print(f"  - Tep da bi xoa (Deleted)   : {status['deleted_count']} tep")
        print("-" * 65)
        print(f"{BOLD}DANH SACH DON TAN CONG DOI KHANG (HARRY HALPIN 2025):{RESET}")
        print(f"  {RED}[1]{RESET} Pha hoai / Sua doi noi dung tep (In-Place File Tamper)")
        print(f"  {RED}[2]{RESET} Chen cua sau / Tep la vao thu muc (Inject Backdoor)")
        print(f"  {RED}[3]{RESET} Xoa len tep tin du lieu (Silent File Deletion)")
        print(f"  {RED}[4]{RESET} Be gay ma tuyen tinh CRC32 bang bu GF(2) (CRC32 Linear Forge)")
        print(f"  {RED}[5]{RESET} Trao doi Merkle Proof tren mang Dolev-Yao (Cross-Proof Swap)")
        print(f"  {RED}[6]{RESET} Tan cong Tien anh cap hai (Second-Preimage Attack)")
        print(f"  {RED}[7]{RESET} Do dac kenh phu thoi gian (Kocher 1996 Timing Side-Channel)")
        print(f"  {RED}[8]{RESET} Thu thap sieu du lieu & Luu luong (Traffic & Metadata Recon)")
        print(f"  {GREEN}[9] Khoi phuc nguyen trang (Rollback All){RESET}")
        print(f"  {CYAN}[10] Xem nhat ky tan cong (Attack Logs){RESET}")
        print(f"  {YELLOW}[0] Thoat{RESET}")

        choice = input(f"\n{BOLD}RedTeam > {RESET}").strip()

        if choice == "1":
            print(f"\n[*] {YELLOW}Dang phat dong tan cong sua doi tep tin...{RESET}")
            res = engine.attack_tamper_file()
            if res.get("success"):
                print(f"{GREEN}[+] TAN CONG THANH CONG!{RESET}")
                print(f"    - Tep bi pha hoai  : {BOLD}{res['file']}{RESET}")
                print(f"    - SHA-256 goc      : {res['original_sha256']}")
                print(f"    - SHA-256 sau sua  : {RED}{res['tampered_sha256']}{RESET}")
                print(f"    - Phan ung BlueTeam: {YELLOW}{res['expected_blue_team_reaction']}{RESET}")
            else:
                print(f"{RED}[-] That bai: {res.get('error')}{RESET}")

        elif choice == "2":
            fname = input("Nhap ten tep backdoor muon tha (Enter de mac dinh): ").strip()
            res = engine.attack_inject_backdoor(fname if fname else None)
            if res.get("success"):
                print(f"{GREEN}[+] THA CUA SAU THANH CONG!{RESET}")
                print(f"    - Tep la xam nhap  : {MAGENTA}{res['injected_file']}{RESET}")
                print(f"    - Kich thuoc       : {res['file_size']} Bytes")
                print(f"    - SHA-256          : {res['sha256']}")
                print(f"    - Phan ung BlueTeam: {YELLOW}{res['expected_blue_team_reaction']}{RESET}")

        elif choice == "3":
            print(f"\n[*] {YELLOW}Dang thuc hien xoa len tep tin khoi dia...{RESET}")
            res = engine.attack_delete_file()
            if res.get("success"):
                print(f"{GREEN}[+] XOA TEP THANH CONG!{RESET}")
                print(f"    - Tep da xoa       : {RED}{res['deleted_file']}{RESET}")
                print(f"    - Phan ung BlueTeam: {YELLOW}{res['expected_blue_team_reaction']}{RESET}")

        elif choice == "4":
            c_msg = input("Nhap thong diep gia mao (Enter de dung thong diep mau): ").strip()
            print(f"\n[*] {YELLOW}Dang giai ma tran GF(2) bang thuat toan Khu Gauss de sinh 4 byte bu...{RESET}")
            res = engine.attack_crc32_forge(c_msg if c_msg else None)
            if res.get("success"):
                print(f"{GREEN}[+] BE GAY MA TUYEN TINH CRC32 THANH CONG!{RESET}")
                print(f"    - Thong diep goc   : {res['original_message']}")
                print(f"    - Thong diep gia   : {RED}{res['tampered_message']}{RESET}")
                print(f"    - 4 Byte bu dai so : {BOLD}{CYAN}{res['algebraic_patch_hex']}{RESET}")
                print(f"    - Ket qua CRC32    : Goc {res['crc32_result']['original_crc']} == Sau sua {res['crc32_result']['forged_crc']}")
                print(f"    - Trang thai CRC32 : {RED}{BOLD}{res['crc32_result']['crc32_status']}{RESET}")
                print(f"    - Trang thai SHA   : {GREEN}{BOLD}{res['sha256_result']['sha256_status']}{RESET}")
                print(f"    - Ket luan hoc thuat: {res['academic_conclusion']}")

        elif choice == "5":
            print(f"\n[*] {YELLOW}Dang gia lap Dolev-Yao chan duong truyen mang va trao doi Proof...{RESET}")
            res = engine.attack_dolev_yao_swap()
            if res.get("success"):
                print(f"{GREEN}[+] THUC HIEN TRAO DOI PROOF XONG!{RESET}")
                print(f"    - Hanh vi ke thu   : {res['attacker_action']}")
                print(f"    - So buoc proof    : {res['proof_steps']} steps")
                print(f"    - Blue Team xu ly  : {GREEN}{BOLD}{res['blue_team_verification']}{RESET}")

        elif choice == "6":
            print(f"\n[*] {YELLOW}Dang che tao nut la gia mao nhi phan (Second-Preimage Attack)...{RESET}")
            res = engine.attack_second_preimage_exploit()
            if res.get("success"):
                print(f"{GREEN}[+] KET QUA DOI SANH TIEN ANH THU HAI:{RESET}")
                print(f"    - Standard Merkle  : {RED}{BOLD}{res['standard_merkle_tree']['attack_result']}{RESET}")
                print(f"    - HD-Merkle (RFC)  : {GREEN}{BOLD}{res['hd_merkle_tree']['attack_result']}{RESET}")

        elif choice == "7":
            print(f"\n[*] {YELLOW}Dang do dac kenh phu thoi gian (Kocher 1996 Timing Side-Channel)...{RESET}")
            res = engine.attack_timing_side_channel()
            if res.get("success"):
                print(f"{GREEN}[+] KET QUA DO DAC KENH PHU THOI GIAN:{RESET}")
                print(f"    - Early-exit mismatch byte 1: {res['naive_comparison']['early_exit_mismatch_first_byte_us']} us")
                print(f"    - Late-exit mismatch byte 64: {res['naive_comparison']['late_exit_mismatch_last_byte_us']} us")
                print(f"    - Ti le ro ri thoi gian     : {RED}{BOLD}{res['naive_comparison']['timing_leak_ratio']}x{RESET}")
                print(f"    - Constant-time HMAC defense: {GREEN}{BOLD}{res['constant_time_defense']['status']}{RESET}")

        elif choice == "8":
            print(f"\n[*] {YELLOW}Dang thu thap sieu du lieu & luu luong (Traffic & Metadata Analysis)...{RESET}")
            res = engine.attack_metadata_recon()
            if res.get("success"):
                print(f"{GREEN}[+] TRICH XUAT SIEU DU LIEU HOAN TAT!{RESET}")
                info = res["traffic_metadata_intelligence"]
                print(f"    - Tong so tep giam sat : {info['total_monitored_files']} tep")
                print(f"    - Tong dung luong      : {info['total_size_mb']} MB")
                print(f"    - Thu muc con hang dau : {info['top_subdirectories']}")
                print(f"    - Suy luan ke tan cong : {YELLOW}{res['adversary_deduction']}{RESET}")

        elif choice == "9":
            print(f"\n[*] {YELLOW}Dang khoi phuc toan bo tep ve trang thai sach ban dau...{RESET}")
            res = engine.rollback_all()
            print(f"{GREEN}[+] {res['status']}{RESET}")
            print(f"    - Phuc hoi tep sua : {res['restored_tampered_files']} tep")
            print(f"    - Don dep tep la   : {res['cleaned_injected_files']} tep")
            print(f"    - Phuc hoi tep xoa : {res['restored_deleted_files']} tep")

        elif choice == "10":
            print(f"\n{CYAN}{BOLD}=== LICH SU CAC DON TAN CONG (LATEST 10 EVENTS) ==={RESET}")
            logs = engine.get_status().get("latest_logs", [])
            if not logs:
                print("Chua co su kien tan cong nao duoc ghi nhan.")
            for log in logs:
                print(f"[{log['timestamp']}] {BOLD}{log['attack']}{RESET} -> {log['status']}")
            print("=" * 65)

        elif choice == "0":
            print(f"\n{YELLOW}[*] Dang thoat Red Team Console. Tam biet!{RESET}")
            break
        else:
            print(f"{RED}[!] Lua chon khong hop le. Vui long chon lai.{RESET}")

        print("\n" + "=" * 65 + "\n")


if __name__ == "__main__":
    main()
