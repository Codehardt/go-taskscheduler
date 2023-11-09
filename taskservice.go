package taskscheduler

import (
	"errors"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// Task is a task found in Windows Task Scheduler 2.0
type Task struct {
	Name        string
	Path        string
	Enabled     bool
	LastRunTime time.Time
	NextRunTime time.Time
	ActionList  []ExecAction // Other actions are ignored, we are only interested in Commandline Actions
}

// ExecAction is an action defined in a scheduled Task if type IExecAction.
type ExecAction struct {
	WorkingDirectory string
	Path             string
	Arguments        string
}

// GetTasks returns a list of all scheduled Tasks in Windows Task Scheduler 2.0
func GetTasks() ([]Task, error) {
	// Initialize COM API
	if err := ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED); err != nil {
		return nil, errors.New("Could not initialize Windows COM API")
	}
	defer ole.CoUninitialize()
	// Create an ITaskService object
	unknown, err := ole.CreateInstance(ole.NewGUID("{0F87369F-A4E5-4CFC-BD3E-73E6154572DD}"), nil)
	if err != nil {
		return nil, errors.New("Could not initialize Task Scheduler 2.0")
	}
	defer unknown.Release()
	// Convert IUnknown to IDispatch to get more functions like CallMethod()
	ts, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil, errors.New("Could not prepare Task Scheduler 2.0")
	}
	defer ts.Release()
	// Connect to the Task Scheduler 2.0
	if _, err := ts.CallMethod("Connect", "", "", "", ""); err != nil {
		return nil, errors.New("Could not connect to Task Scheduler 2.0")
	}
	// Get Root Directory of Task Scheduler 2.0 and get all tasks recursively
	variant, err := oleutil.CallMethod(ts, "GetFolder", "\\")
	if err != nil {
		return nil, errors.New("Could not get root folder in Task Scheduler 2.0")
	}
	root := variant.ToIDispatch()
	defer root.Release()
	return getTasksRecursively(root), nil
}

func getTasksRecursively(folder *ole.IDispatch) (tasks []Task) {
	var (
		variant *ole.VARIANT
		err     error
	)
	// Get Tasks in subfolders first
	if variant, err = oleutil.CallMethod(folder, "GetFolders", int64(0)); err != nil {
		return
	}
	folderIterator := variant.ToIDispatch()
	if variant, err = oleutil.GetProperty(folderIterator, "count"); err != nil {
		return
	}
	count, _ := variant.Value().(int32)
	for i := int32(1); i <= count; i++ {
		// Get Tasks of subfolder i
		index := ole.NewVariant(ole.VT_I4, int64(i))
		if variant, err = oleutil.GetProperty(folderIterator, "item", &index); err != nil {
			continue
		}
		subfolder := variant.ToIDispatch()
		subtasks := getTasksRecursively(subfolder)
		tasks = append(tasks, subtasks...)
		subfolder.Release()
	}
	folderIterator.Release()
	// Get Tasks
	if variant, err = oleutil.CallMethod(folder, "GetTasks", int64(0)); err != nil {
		return
	}
	taskIterator := variant.ToIDispatch()
	if variant, err = oleutil.GetProperty(taskIterator, "count"); err != nil {
		return
	}
	count, _ = variant.Value().(int32)
	for i := int32(1); i <= count; i++ {
		// Get Task i
		index := ole.NewVariant(ole.VT_I4, int64(i))
		if variant, err = oleutil.GetProperty(taskIterator, "item", &index); err != nil {
			continue
		}
		task := variant.ToIDispatch()
		var t Task
		if variant, err = oleutil.GetProperty(task, "name"); err == nil {
			t.Name = variant.ToString()
		}
		if variant, err = oleutil.GetProperty(task, "path"); err == nil {
			t.Path = variant.ToString()
		}
		if variant, err = oleutil.GetProperty(task, "enabled"); err == nil {
			t.Enabled, _ = variant.Value().(bool)
		}
		if variant, err = oleutil.GetProperty(task, "lastRunTime"); err == nil {
			t.LastRunTime, _ = variant.Value().(time.Time)
		}
		if variant, err = oleutil.GetProperty(task, "nextRunTime"); err == nil {
			t.NextRunTime, _ = variant.Value().(time.Time)
		}
		// Get more details, e.g. actions
		if variant, err = oleutil.GetProperty(task, "definition"); err == nil {
			definition := variant.ToIDispatch()
			if variant, err = oleutil.GetProperty(definition, "actions"); err == nil {
				actions := variant.ToIDispatch()
				if variant, err = oleutil.GetProperty(actions, "count"); err == nil {
					count2, _ := variant.Value().(int32)
					for i := int32(1); i <= count2; i++ {
						// Get Action i
						index := ole.NewVariant(ole.VT_I4, int64(i))
						if variant, err = oleutil.GetProperty(actions, "item", &index); err != nil {
							continue
						}
						action := variant.ToIDispatch()
						if variant, err = oleutil.GetProperty(action, "type"); err != nil {
							action.Release()
							continue
						}
						actionType, _ := variant.Value().(int32)
						if actionType != 0 { // only handle IExecAction
							action.Release()
							continue
						}
						var a ExecAction
						if variant, err = oleutil.GetProperty(action, "workingDirectory"); err == nil {
							a.WorkingDirectory = variant.ToString()
						}
						if variant, err = oleutil.GetProperty(action, "path"); err == nil {
							a.Path = variant.ToString()
						}
						if variant, err = oleutil.GetProperty(action, "arguments"); err == nil {
							a.Arguments = variant.ToString()
						}
						t.ActionList = append(t.ActionList, a)
						action.Release()
					}
				}
			}
		}
		tasks = append(tasks, t)
		task.Release()
	}
	taskIterator.Release()
	return
}
