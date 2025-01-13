"""전체 코드"""
import cv2
import pyrealsense2 as rs
import numpy as np
import time

# RealSense 카메라 설정
pipeline = rs.pipeline()
config = rs.config()
config.enable_stream(rs.stream.color, 1280, 720, rs.format.bgr8, 30)
config.enable_stream(rs.stream.depth, 1280, 720, rs.format.z16, 30)

# 파이프라인 시작 및 depth_sensor 획득
profile = pipeline.start(config)
depth_sensor = profile.get_device().first_depth_sensor()

# Depth sensor 초기화 확인
if depth_sensor is None:
    print("Depth sensor initialization failed.")
else:
    print("Depth sensor initialized successfully.")

# 레이저 파워 설정
if depth_sensor.supports(rs.option.laser_power):
    max_power = depth_sensor.get_option_range(rs.option.laser_power).max
    depth_sensor.set_option(rs.option.laser_power, max_power)

# IR 게인 초기 설정
if depth_sensor.supports(rs.option.gain):
    initial_gain = 100
    depth_sensor.set_option(rs.option.gain, initial_gain)

# 좌표 저장용 변수
horizontal_points = []
vertical_points = []
mode = "horizontal"  # 현재 좌표 입력 모드: "horizontal" 또는 "vertical"
#depth_sensor = None  # depth_sensor 전역 변수 추가


def get_camera_intrinsics(pipeline):
    """카메라 내부 파라미터 얻기"""
    # 현재 활성화된 프로파일에서 color 스트림 가져오기
    profile = pipeline.get_active_profile()
    color_stream = profile.get_stream(rs.stream.color)
    intrinsics = color_stream.as_video_stream_profile().get_intrinsics()
    
    return intrinsics

def optimize_depth_settings(x, y, pipeline, depth_sensor):
    """깊이 측정을 위한 IR 게인과 노출 시간 최적화"""
    
    best_exposure = None
    best_gain = None
    best_quality = 0
    best_depth = 0
    
    # 노출 시간 범위 (마이크로초)
    exposures = [100, 250, 500, 750, 1000]
    # 게인 값 범위 (16-248)
    gains = [16, 32, 64, 128, 248]
    
    try:
        for exposure in exposures:
            for gain in gains:
                # 노출 시간 설정
                depth_sensor.set_option(rs.option.exposure, exposure)
                # IR 게인 설정
                if depth_sensor.supports(rs.option.gain):
                    depth_sensor.set_option(rs.option.gain, gain)
                
                time.sleep(0.1)
                
                depths = []
                for _ in range(10):
                    frames = pipeline.wait_for_frames()
                    depth_frame = frames.get_depth_frame()
                    depth = depth_frame.get_distance(x, y)
                    if depth > 0:
                        depths.append(depth)
                
                if depths:
                    mean_depth = np.mean(depths)
                    std_depth = np.std(depths)
                    quality = 1 / (std_depth + 1e-6)
                    
                    print(f"Testing - Exposure: {exposure}μs, Gain: {gain}, "
                          f"Depth: {mean_depth:.3f}m, Quality: {quality:.2f}")
                    
                    if quality > best_quality:
                        best_quality = quality
                        best_exposure = exposure
                        best_gain = gain
                        best_depth = mean_depth
        
        if best_exposure and best_gain:
            depth_sensor.set_option(rs.option.exposure, best_exposure)
            if depth_sensor.supports(rs.option.gain):
                depth_sensor.set_option(rs.option.gain, best_gain)
            
            print(f"\nOptimal settings found:")
            print(f"Exposure: {best_exposure}μs")
            print(f"Gain: {best_gain}")
            print(f"Depth stability: ±{1/best_quality:.4f}m")
            
        return best_depth
        
    except Exception as e:
        print(f"Error during optimization: {e}")
        return None

# 클릭 콜백 함수
def select_point(event, x, y, flags, param):
    global horizontal_points, vertical_points, mode, pipeline, depth_sensor
    
    if event == cv2.EVENT_LBUTTONDOWN:
        if depth_sensor is None:  # depth_sensor가 None인지 확인
            print("Depth sensor is not initialized.")
            return
        # 선택한 점에 대해 노출 최적화
        optimized_depth = optimize_depth_settings(x, y, pipeline, depth_sensor)
        
        if optimized_depth > 0:
            if mode == "horizontal" and len(horizontal_points) < 2:
                horizontal_points.append((x, y))
                print(f"Horizontal point {len(horizontal_points)} selected at depth: {optimized_depth:.3f}m")
                
                if len(horizontal_points) == 2:
                    # intrinsics 획득 및 거리 계산
                    intrinsics = get_camera_intrinsics(pipeline)
                    frames = pipeline.wait_for_frames()
                    depth_frame = frames.get_depth_frame()
                    dist = calculate_distance_3d(horizontal_points[0], horizontal_points[1], depth_frame, intrinsics)
                    print(f"Horizontal distance: {dist:.2f}mm")
                    print("수평 좌표 입력이 완료되었습니다. 수직에 대한 좌표 두 개를 찍어주세요.")
                    mode = "vertical"
            
            elif mode == "vertical" and len(vertical_points) < 2:
                vertical_points.append((x, y))
                print(f"Vertical point {len(vertical_points)} selected at depth: {optimized_depth:.3f}m")
                
                if len(vertical_points) == 2:
                    intrinsics = get_camera_intrinsics(pipeline)
                    frames = pipeline.wait_for_frames()
                    depth_frame = frames.get_depth_frame()
                    dist = calculate_distance_3d(vertical_points[0], vertical_points[1], depth_frame, intrinsics)
                    print(f"Vertical distance: {dist:.2f}mm")
                    print("수직 좌표 입력이 완료되었습니다.")

def calculate_distance_3d(p1, p2, depth_frame, intrinsics):
    """두 점의 3D 거리 계산 (깊이 포함)"""
    # 3D 좌표로 변환
    p1_depth = depth_frame.get_distance(p1[0], p1[1])
    p2_depth = depth_frame.get_distance(p2[0], p2[1])
    # print(f"p1:  {p1_depth}")
    # print(f"p2:  {p2_depth}")

    # 픽셀 좌표를 3D 공간 좌표로 변환
    p1_3d = rs.rs2_deproject_pixel_to_point(intrinsics, [p1[0], p1[1]], p1_depth)
    p2_3d = rs.rs2_deproject_pixel_to_point(intrinsics, [p2[0], p2[1]], p2_depth)
    

    # 3D 공간에서의 유클리드 거리 계산 (mm 단위)
    distance = np.sqrt(
        (p2_3d[0] - p1_3d[0])**2 + 
        (p2_3d[1] - p1_3d[1])**2 + 
        (p2_3d[2] - p1_3d[2])**2
    ) * 1000
    
    # 디버깅을 위한 출력
    print(f"Point 1: Depth = {p1_depth:.3f}m, 3D coords = {p1_3d}")
    print(f"Point 2: Depth = {p2_depth:.3f}m, 3D coords = {p2_3d}")
    print(f"Calculated 3D distance: {distance:.2f} mm")
    return distance

# 윈도우 생성 및 클릭 이벤트 등록
cv2.namedWindow("RealSense")
cv2.setMouseCallback("RealSense", select_point)

print("수평 좌표를 두개 찍어주세요. ")

try:
    while True:
        # 프레임 가져오기
        frames = pipeline.wait_for_frames()
        color_frame = frames.get_color_frame()
        depth_frame = frames.get_depth_frame()

        if not color_frame or not depth_frame:
            continue

        # 이미지를 NumPy 배열로 변환
        color_image = np.asanyarray(color_frame.get_data())
        depth_scale = pipeline.get_active_profile().get_device().first_depth_sensor().get_depth_scale()

        # 수평, 수직 선 표시
        if len(horizontal_points) == 2:
            cv2.line(color_image, horizontal_points[0], horizontal_points[1], (0, 0, 255), 2)  # 빨간색
        if len(vertical_points) == 2:
            cv2.line(color_image, vertical_points[0], vertical_points[1], (255, 0, 0), 2)  # 파란색

        # 키 입력 대기
        key = cv2.waitKey(1)
        if key == ord('a'):  # 계산 수행
            if len(horizontal_points) == 2 and len(vertical_points) == 2:
                horizontal_length = calculate_distance_3d(horizontal_points[0], horizontal_points[1], depth_frame, get_camera_intrinsics(pipeline))
                vertical_length = calculate_distance_3d(vertical_points[0], vertical_points[1], depth_frame, get_camera_intrinsics(pipeline))

                # 결과 표시
                cv2.putText(color_image, f"Horizontal: {horizontal_length:.2f} mm", (10, 400), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 0, 255), 2)
                cv2.putText(color_image, f"Vertical: {vertical_length:.2f} mm", (10, 430), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 0, 0), 2)
                print(f"Horizontal Distance: {horizontal_length:.2f} mm")
                print(f"Vertical Distance: {vertical_length:.2f} mm")
                
            else:
                print("모든 좌표를 입력해야 계산할 수 있습니다.")

        elif key == ord('q'):  # 종료
            break

        # 이미지 표시
        cv2.imshow("RealSense", color_image)

finally:
    pipeline.stop()
    cv2.destroyAllWindows()
