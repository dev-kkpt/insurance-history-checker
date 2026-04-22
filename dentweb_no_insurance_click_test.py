import argparse
import time

from test3_reconstructed import (
    check_for_verification_window,
    find_EditText,
    find_and_maximize_dentweb,
    get_patient_list,
    input_text_to_field,
    show_completion_popup,
)


def run_test(limit):
    app, window = find_and_maximize_dentweb()
    if not window:
        print("DentWeb 창을 찾지 못했습니다.")
        return

    name_field = find_EditText(window)
    if not name_field:
        print("검색 입력창을 찾지 못했습니다.")
        return

    patient_list, _ = get_patient_list()
    if not patient_list:
        print("조회 대상 환자가 없습니다.")
        return

    if limit is not None:
        patient_list = patient_list[:limit]

    print(f"총 {len(patient_list)}명의 환자에 대해 안전 테스트를 진행합니다.")
    print("보험자격조회 버튼 클릭은 생략합니다.")

    for patient in patient_list:
        (
            _patient_id,
            chart_no,
            _patient_type,
            patient_name,
            *_rest,
        ) = patient
        search_text = f"{patient_name}({chart_no})"
        print(f"\n[TEST] 환자 검색: {search_text}")
        input_text_to_field(name_field, search_text)
        check_for_verification_window(app)
        print("[SKIP] 보험자격조회 버튼 클릭 생략")
        time.sleep(1)

    show_completion_popup()


def main():
    parser = argparse.ArgumentParser(
        description="현재 운영 흐름에서 보험자격조회 클릭만 제외하고 안전하게 테스트합니다."
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=5,
        help="테스트할 환자 수입니다. 기본값은 5명입니다.",
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="조건에 맞는 전체 환자를 대상으로 테스트합니다.",
    )
    args = parser.parse_args()

    limit = None if args.all else args.limit
    run_test(limit)


if __name__ == "__main__":
    main()
