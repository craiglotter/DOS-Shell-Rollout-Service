DOS-Shell-Rollout-Service
=========================

DOS Shell Rollout Service is a Windows service that runs in the background to ensure that the DOS Shell Rollout program is actually running, in an attempt to stop network PC users from closing the remote update application down. It simply monitors the Windows processes and if it is discovered that DOS Shell Rollout is not running, it submits an entry to the Invisible Application Starter's Monitor directory containing DOS Shell Rollout's executable path. Created by Craig Lotter, December 2005
