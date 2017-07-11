# -*- coding: utf-8 -*-



import os, sys
import time



import argparse
import datetime
import imutils
import time
import cv2


g_miniArea = 150
g_grabValue = 80
# g_updateEclipse = 24*60*30
g_updateEclipse = 10

def ParseCmdLine():
    # 创建参数解析器并解析参数
    ap = argparse.ArgumentParser()
    ap.add_argument("-v", "--video", help="path to the video file")
    ap.add_argument("-a", "--min-area", type=int, default=500, help="minimum area size")
    args = vars(ap.parse_args())
    # 如果video参数为None，那么我们从摄像头读取数据
    if args.get("video", None) is None:
        camera = cv2.VideoCapture(0)
        time.sleep(0.25)



def ParseVideo(tFile):

    if os.path.exists(tFile):
        camera = cv2.VideoCapture(tFile)
        print u"load file into opencv %s"%tFile
    else:
        print u"file is not exist %s"%tFile
        return
    # 初始化视频流的第一帧
    dWidth = int(camera.get(cv2.CAP_PROP_FRAME_WIDTH ))
    dHeight =int(camera.get(cv2.CAP_PROP_FRAME_HEIGHT))
    firstFrame = None
    lIndex = 0
    fourcc = cv2.VideoWriter_fourcc(*'XVID')
    lFileName = str.format("output_%s.avi"%tFile)
    lOutPutWriter = cv2.VideoWriter(lFileName,fourcc, 20.0, (dWidth,dHeight))

    while True:
        # 获取当前帧并初始化occupied/unoccupied文本
        (grabbed, frame) = camera.read()
        lIndex += 1
        print "\r",
        print "read frame index  %d" % (lIndex),
        # 如果不能抓取到一帧，说明我们到了视频的结尾
        if not grabbed:
            print u"end of video"
            break

        # 调整该帧的大小，转换为灰阶图像并且对其进行高斯模糊
        frame = imutils.resize(frame, width=dWidth,height=dHeight)
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        gray = cv2.GaussianBlur(gray, (21, 21), 0)

        # 如果第一帧是None，对其进行初始化
        if firstFrame is None:
            firstFrame = gray
            print u"第一帧是None，对其进行初始化"
            continue
        # 每个1小时自动更新对比图片
        if lIndex%g_updateEclipse == 0:
            firstFrame = gray
            print u"update bkimage %d"%lIndex

        # 计算当前帧和第一帧的不同
        frameDelta = cv2.absdiff(firstFrame, gray)
        thresh = cv2.threshold(frameDelta, g_grabValue, 255, cv2.THRESH_BINARY)[1]

        # 扩展阀值图像填充孔洞，然后找到阀值图像上的轮廓
        thresh = cv2.dilate(thresh, None, iterations=2)
        (image, contours, hierarchy ) = cv2.findContours(thresh.copy(), cv2.RETR_EXTERNAL,
                                     cv2.CHAIN_APPROX_SIMPLE)

        # 遍历轮廓
        for c in contours:
            # if the contour is too small, ignore it
            area = cv2.contourArea(c)
            if area < g_miniArea:
                # print "\r",
                # print u"ignore contour as too small skip: %d "%area
                continue
            # compute the bounding box for the contour, draw it on the frame,
            # and update the text
            # 计算轮廓的边界框，在当前帧中画出该框
            (x, y, w, h) = cv2.boundingRect(c)
            cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)
            text = "index: %d page: %d diff: %7d  "%(lIndex,lIndex/g_updateEclipse,area,)
            # draw the text and timestamp on the frame
            # 在当前帧上写文字以及时间戳
            cv2.putText(frame, "{}".format(text), (10, 15),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 0, 255), 2)
            # cv2.putText(frame, datetime.datetime.now().strftime("%A %d %B %Y %I:%M:%S%p"),
            #             (10, frame.shape[0] - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.35, (0, 0, 255), 1)

            lOutPutWriter.write(frame)
            #显示当前帧并记录用户是否按下按键
            # cv2.imshow("Security Feed", frame)
            # # cv2.imshow("Thresh", thresh)
            # cv2.imshow("Frame Delta", frameDelta)

            print u"record index %d"%lIndex
            # 如果q键被按下，跳出循环
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break

    # 清理摄像机资源并关闭打开的窗口
    camera.release()
    lOutPutWriter.release()
    cv2.destroyAllWindows()
    print u"done release source"


def TestPlayFile():
    import numpy as np
    import cv2
    cap = cv2.VideoCapture('test.mp4')
    while(cap.isOpened()):
        ret, frame = cap.read()
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        cv2.imshow('frame',gray)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
    cap.release()
    cv2.destroyAllWindows()
    print u"TestPlayFile done"


def TestPlayCamera():
    import numpy as np
    import cv2
    cap = cv2.VideoCapture(0)
    while(True):
        # Capture frame-by-frame
        ret, frame = cap.read()

        # Our operations on the frame come here
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

        # Display the resulting frame
        cv2.imshow('frame',gray)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    # When everything done, release the capture
    cap.release()
    cv2.destroyAllWindows()

def TestSaveVideo():
    import numpy as np
    import cv2
    cap = cv2.VideoCapture(0)
    # Define the codec and create VideoWriter object
    fourcc = cv2.VideoWriter_fourcc(*'XVID')
    out = cv2.VideoWriter('output.mp4',fourcc, 20.0, (640,480))

    while(cap.isOpened()):
        ret, frame = cap.read()
        if ret==True:
            frame = cv2.flip(frame,0)
            # write the flipped frame
            out.write(frame)
            cv2.imshow('frame',frame)
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break
        else:
            break

    # Release everything if job is finished
    cap.release()
    out.release()
    cv2.destroyAllWindows()


if __name__ == "__main__":
    lFile = "test3.mp4"
    ParseVideo(lFile)
    # TestPlayFile()
    # TestPlayCamera()
    # TestSaveVideo()